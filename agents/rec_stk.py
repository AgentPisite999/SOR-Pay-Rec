

# import os
# import re
# import json
# import base64
# import math
# import argparse
# from pathlib import Path

# import pandas as pd
# from dotenv import load_dotenv
# import gspread
# from google.oauth2.service_account import Credentials

# # ===================== SHEET NAMES =====================
# DEST_SHEET_NAME = "Rec_STK_Base"
# PIVOT_SHEET_NAME = "Rec_Stk_Final"

# # ===================== CLI =====================
# def parse_args():
#     ap = argparse.ArgumentParser(description="Receivable Stock (rec_stk) processor")
#     ap.add_argument("--input_dir", required=True, help="Folder containing input files")
#     ap.add_argument("--output_dir", required=True, help="Directory to write output Excel")
#     ap.add_argument(
#         "--output_name",
#         default="REC_STK_MASTER.xlsx",
#         help="Output Excel file name (default: REC_STK_MASTER.xlsx)",
#     )
#     ap.add_argument("--month", default="", help="Month label (e.g. JAN, FEB, ...)")
#     ap.add_argument("--year", default="", help="Year (e.g. 2026)")
#     return ap.parse_args()

# # -------------------- .env helper --------------------
# def env_clean(v: str, default: str = "") -> str:
#     if v is None:
#         return default
#     s = str(v).strip()
#     if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
#         s = s[1:-1].strip()
#     if s.endswith(";"):
#         s = s[:-1].strip()
#     return s or default

# # -------------------- Common helpers --------------------
# def canon(x: str) -> str:
#     s = str(x).replace("\u00A0", " ").strip().lower()
#     s = re.sub(r"[_\-\.\(\)\[\]\/\\]+", " ", s)
#     s = " ".join(s.split())
#     return s

# def build_canon_map(cols):
#     m = {}
#     for c in cols:
#         m[canon(c)] = c
#     return m

# def pick_first_existing(col_map, candidates):
#     for c in candidates:
#         if c in col_map:
#             return col_map[c]
#     return None

# def to_num(s):
#     return pd.to_numeric(s, errors="coerce")

# def normalize_ean(x) -> str:
#     if x is None:
#         return ""
#     if isinstance(x, int):
#         return str(x)
#     if isinstance(x, float):
#         if math.isfinite(x):
#             return str(int(round(x)))
#         return ""
#     s = str(x).replace("\u00A0", " ").strip()
#     s = re.sub(r"\.0$", "", s)
#     if re.match(r"^\d+(\.\d+)?e\+\d+$", s.lower()):
#         try:
#             return str(int(round(float(s))))
#         except:
#             return s
#     return s

# def infer_customer_from_filename(filename: str) -> str:
#     f = canon(filename)
#     if "flipkart" in f:
#         return "Flipkart"
#     if "myntra" in f:
#         return "Myntra"
#     if "ajio" in f:
#         return "Reliance Ajio SOR"
#     return ""

# def drop_fully_blank_rows(df4: pd.DataFrame) -> pd.DataFrame:
#     df4 = df4.copy()
#     df4["Customer Name"] = df4["Customer Name"].astype(str).fillna("").str.strip()
#     df4["Barcode"] = df4["Barcode"].astype(str).fillna("").str.strip()

#     mask_any = (
#         df4["Customer Name"].ne("")
#         | df4["Barcode"].ne("")
#         | df4["Qty."].notna()
#         | df4["MRP Rate"].notna()
#     )
#     return df4[mask_any]

# # -------------------- Read files (all sheets) --------------------
# def read_any_file_all_sheets(path: Path):
#     suf = path.suffix.lower()

#     if suf == ".csv":
#         try:
#             return [pd.read_csv(path)]
#         except UnicodeDecodeError:
#             return [pd.read_csv(path, encoding="latin1")]

#     if suf in (".xlsx", ".xls"):
#         xl = pd.ExcelFile(path)
#         return [pd.read_excel(path, sheet_name=sn) for sn in xl.sheet_names]

#     if suf == ".xlsb":
#         xl = pd.ExcelFile(path, engine="pyxlsb")
#         return [pd.read_excel(path, sheet_name=sn, engine="pyxlsb") for sn in xl.sheet_names]

#     raise ValueError(f"Unsupported file type: {path.name}")

# # -------------------- 1) Build base Rec_STK_Base --------------------
# def ensure_base_cols(df: pd.DataFrame, file_path: Path) -> pd.DataFrame:
#     df = df.copy()
#     df.columns = [str(c).replace("\u00A0", " ").strip() for c in df.columns]
#     col_map = build_canon_map(df.columns)

#     barcode_col = pick_first_existing(col_map, ["eancode", "ean code", "ean", "barcode", "bar code"])
#     mrp_col = pick_first_existing(col_map, ["mrp rate", "mrp"])

#     fixed = infer_customer_from_filename(file_path.name)

#     if fixed == "Flipkart":
#         qty_col = pick_first_existing(col_map, ["soh", "soh qty"])
#         customer_series = pd.Series(["Flipkart"] * len(df))

#     elif fixed == "Myntra":
#         qty_col = pick_first_existing(col_map, ["count_", "count"])
#         customer_series = pd.Series(["Myntra"] * len(df))

#     elif fixed == "Reliance Ajio SOR":
#         qty_col = pick_first_existing(col_map, ["count_", "count", "active qty", "active_qty"])
#         customer_series = pd.Series(["Reliance Ajio SOR"] * len(df))

#     else:
#         partner_col = pick_first_existing(col_map, ["partner", "partner name"])
#         qty_col = pick_first_existing(col_map, ["closing stock", "closing_stock"])
#         customer_series = (
#             df[partner_col].astype(str).fillna("").str.strip()
#             if partner_col else pd.Series([""] * len(df))
#         )

#     barcode_series = df[barcode_col].map(normalize_ean) if barcode_col else pd.Series([""] * len(df))
#     qty_series = to_num(df[qty_col]) if qty_col else pd.Series([pd.NA] * len(df))
#     mrp_series = to_num(df[mrp_col]) if mrp_col else pd.Series([pd.NA] * len(df))

#     out = pd.DataFrame({
#         "Customer Name": customer_series,
#         "Barcode": barcode_series,
#         "Qty.": qty_series,
#         "MRP Rate": mrp_series,
#     })
#     return out

# def add_month_year_cols(df: pd.DataFrame, month: str, year: str) -> pd.DataFrame:
#     df = df.copy()
#     month = (month or "").strip().upper()
#     year = (year or "").strip()
#     if month and year:
#         df.insert(0, "Month", month)
#         df.insert(0, "Year", year)
#     return df

# # -------------------- Google Sheets --------------------
# def get_gspread_client_from_env(write: bool = False):
#     load_dotenv()
#     b64 = os.getenv("GOOGLE_SA_JSON_B64")
#     if not b64:
#         raise RuntimeError("Missing GOOGLE_SA_JSON_B64 in .env (must be one-line base64).")
#     sa_json = json.loads(base64.b64decode(b64).decode("utf-8"))
#     scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     if write:
#         scopes = ["https://www.googleapis.com/auth/spreadsheets"]
#     creds = Credentials.from_service_account_info(sa_json, scopes=scopes)
#     return gspread.authorize(creds)

# def worksheet_to_df(ws) -> pd.DataFrame:
#     vals = ws.get_all_values()
#     if not vals:
#         return pd.DataFrame()
#     headers = vals[0]
#     rows = vals[1:]
#     df = pd.DataFrame(rows, columns=headers)
#     df.columns = [str(c).strip() for c in df.columns]
#     return df

# def _normalize_month_year_for_compare(df: pd.DataFrame):
#     if "Year" in df.columns:
#         df["Year"] = df["Year"].astype(str).str.strip()
#     if "Month" in df.columns:
#         df["Month"] = df["Month"].astype(str).str.strip().str.upper()
#     return df

# def _normalize_customer_name_for_compare(df: pd.DataFrame):
#     if "Customer Name" in df.columns:
#         df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
#     return df

# def _align_columns(existing_df: pd.DataFrame, df_new: pd.DataFrame):
#     for c in existing_df.columns:
#         if c not in df_new.columns:
#             df_new[c] = ""
#     for c in df_new.columns:
#         if c not in existing_df.columns:
#             existing_df[c] = ""

#     existing_cols = list(existing_df.columns)
#     extra_cols = [c for c in df_new.columns if c not in existing_cols]
#     final_cols = existing_cols + extra_cols

#     existing_df = existing_df[final_cols]
#     df_new = df_new[final_cols]
#     return existing_df, df_new

# def upsert_append_by_month_year_customer(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
#     """
#     Pivot push logic:
#     - match by Year + Month + Customer Name
#     - matched existing rows are removed
#     - new rows are appended
#     - unrelated rows stay untouched
#     """
#     gc = get_gspread_client_from_env(write=True)
#     sh = gc.open_by_key(sheet_id)

#     try:
#         ws = sh.worksheet(tab_name)
#     except gspread.WorksheetNotFound:
#         ws = sh.add_worksheet(title=tab_name, rows=1000, cols=max(10, new_df.shape[1] + 2))

#     existing_vals = ws.get_all_values()
#     if not existing_vals:
#         existing_df = pd.DataFrame()
#     else:
#         headers = existing_vals[0]
#         rows = existing_vals[1:]
#         existing_df = pd.DataFrame(rows, columns=headers)
#         existing_df.columns = [str(c).strip() for c in existing_df.columns]

#     df_new = new_df.copy().replace([pd.NA, float("inf"), float("-inf")], "")
#     df_new = df_new.where(pd.notnull(df_new), "")

#     if existing_df.empty:
#         out_df = df_new
#     else:
#         existing_df, df_new = _align_columns(existing_df, df_new)

#         existing_df = _normalize_month_year_for_compare(existing_df)
#         df_new = _normalize_month_year_for_compare(df_new)
#         existing_df = _normalize_customer_name_for_compare(existing_df)
#         df_new = _normalize_customer_name_for_compare(df_new)

#         required_cols = {"Year", "Month", "Customer Name"}
#         if required_cols.issubset(existing_df.columns) and required_cols.issubset(df_new.columns):
#             new_keys = set(
#                 zip(
#                     df_new["Year"].astype(str).str.strip(),
#                     df_new["Month"].astype(str).str.upper(),
#                     df_new["Customer Name"].astype(str).str.strip(),
#                 )
#             )

#             existing_keys = list(
#                 zip(
#                     existing_df["Year"].astype(str).str.strip(),
#                     existing_df["Month"].astype(str).str.upper(),
#                     existing_df["Customer Name"].astype(str).str.strip(),
#                 )
#             )

#             mask_same = [key in new_keys for key in existing_keys]
#             existing_df = existing_df.loc[[not matched for matched in mask_same]].copy()

#         out_df = pd.concat([existing_df, df_new], ignore_index=True)

#     if {"Year", "Month", "Customer Name"}.issubset(out_df.columns):
#         out_df["Year"] = out_df["Year"].astype(str).str.strip()
#         out_df["Month"] = out_df["Month"].astype(str).str.strip().str.upper()
#         out_df["Customer Name"] = out_df["Customer Name"].astype(str).str.strip()
#         out_df = out_df.drop_duplicates(subset=["Year", "Month", "Customer Name"], keep="last")

#     out_df = out_df.where(pd.notnull(out_df), "")
#     values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

#     ws.clear()
#     ws.update(values, value_input_option="USER_ENTERED")

# def replace_batch_by_month_year(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
#     """
#     Base push logic:
#     - remove all existing rows for the same Year + Month
#     - append the new batch for that Year + Month
#     - other months remain untouched
#     """
#     month = (month or "").strip().upper()
#     year = (year or "").strip()

#     gc = get_gspread_client_from_env(write=True)
#     sh = gc.open_by_key(sheet_id)

#     try:
#         ws = sh.worksheet(tab_name)
#     except gspread.WorksheetNotFound:
#         ws = sh.add_worksheet(title=tab_name, rows=1000, cols=max(10, new_df.shape[1] + 2))

#     existing_vals = ws.get_all_values()
#     if not existing_vals:
#         existing_df = pd.DataFrame()
#     else:
#         headers = existing_vals[0]
#         rows = existing_vals[1:]
#         existing_df = pd.DataFrame(rows, columns=headers)
#         existing_df.columns = [str(c).strip() for c in existing_df.columns]

#     df_new = new_df.copy().replace([pd.NA, float("inf"), float("-inf")], "")
#     df_new = df_new.where(pd.notnull(df_new), "")

#     if existing_df.empty:
#         out_df = df_new
#     else:
#         existing_df, df_new = _align_columns(existing_df, df_new)

#         existing_df = _normalize_month_year_for_compare(existing_df)
#         df_new = _normalize_month_year_for_compare(df_new)
#         existing_df = _normalize_customer_name_for_compare(existing_df)
#         df_new = _normalize_customer_name_for_compare(df_new)

#         if month and year and {"Year", "Month"}.issubset(existing_df.columns):
#             mask_same_batch = (
#                 (existing_df["Year"].astype(str).str.strip() == year) &
#                 (existing_df["Month"].astype(str).str.upper() == month)
#             )
#             existing_df = existing_df.loc[~mask_same_batch].copy()

#         out_df = pd.concat([existing_df, df_new], ignore_index=True)

#     out_df = out_df.where(pd.notnull(out_df), "")
#     values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

#     ws.clear()
#     ws.update(values, value_input_option="USER_ENTERED")

# # -------------------- 2) Margin rules + COGS map --------------------
# STOP_WORDS = {
#     "PVT", "PRIVATE", "LTD", "LIMITED", "LLP", "INDIA",
#     "INTERNATIONAL", "GLOBAL", "VENTURE", "PVT.", "LTD."
# }

# def norm_name(s: str) -> str:
#     return re.sub(r"[^A-Z0-9]+", " ", str(s or "").upper()).strip()

# def singular_word(w: str) -> str:
#     if len(w) > 3 and w.endswith("S"):
#         return w[:-1]
#     return w

# def tokenize(normed: str):
#     toks = [t.strip() for t in normed.split(" ") if t.strip()]
#     toks = [singular_word(t) for t in toks]
#     toks = [t for t in toks if t and t not in STOP_WORDS]
#     return toks

# def to_percent_number(v):
#     if v is None:
#         return ""
#     if isinstance(v, (int, float)):
#         if pd.isna(v):
#             return ""
#         v = float(v)
#         return v / 100.0 if v > 1 else v
#     s = str(v).strip()
#     if not s:
#         return ""
#     if "%" in s:
#         try:
#             return float(s.replace("%", "").strip()) / 100.0
#         except:
#             return ""
#     try:
#         n = float(s)
#         return n / 100.0 if n > 1 else n
#     except:
#         return ""

# def load_margin_rules_and_cogs_map():
#     load_dotenv()
#     source_spreadsheet_id = env_clean(os.getenv("SOURCE_SPREADSHEET_ID1", ""))
#     margin_sheet_name = env_clean(os.getenv("MARGIN_SHEET_NAME1", "Margin"))
#     cogs_sheet_name = env_clean(os.getenv("COGS_SHEET_NAME1", "COGS"))

#     if not source_spreadsheet_id:
#         raise RuntimeError("Missing SOURCE_SPREADSHEET_ID1 in .env")

#     gc = get_gspread_client_from_env()
#     sh = gc.open_by_key(source_spreadsheet_id)

#     ws_m = sh.worksheet(margin_sheet_name)
#     df_m = worksheet_to_df(ws_m)
#     if df_m.empty:
#         raise RuntimeError("Margin sheet empty")

#     m_map = build_canon_map(df_m.columns)
#     cust_col = pick_first_existing(m_map, ["customer name", "customer"])
#     margin_col = pick_first_existing(m_map, ["margin"])
#     if not cust_col or not margin_col:
#         raise RuntimeError("Margin sheet must have columns: Customer Name, Margin")

#     rules = []
#     for _, row in df_m.iterrows():
#         master_name_raw = str(row.get(cust_col, "")).strip()
#         if not master_name_raw:
#             continue
#         margin_val = to_percent_number(row.get(margin_col, ""))
#         if margin_val == "":
#             continue

#         key_norm = norm_name(master_name_raw.replace("_", " "))
#         toks = tokenize(key_norm)
#         if toks:
#             rules.append({
#                 "tokens": set(toks),
#                 "margin": float(margin_val),
#                 "score": len(toks) * 1000 + len(key_norm),
#             })

#         if "_" in master_name_raw:
#             alias = master_name_raw.split("_")[0]
#             alias_norm = norm_name(alias)
#             alias_toks = tokenize(alias_norm)
#             if alias_toks:
#                 rules.append({
#                     "tokens": set(alias_toks),
#                     "margin": float(margin_val),
#                     "score": len(alias_toks) * 1000 + len(alias_norm),
#                 })

#     rules.sort(key=lambda r: r["score"], reverse=True)

#     ws_c = sh.worksheet(cogs_sheet_name)
#     df_c = worksheet_to_df(ws_c)
#     if df_c.empty:
#         raise RuntimeError("COGS sheet empty")

#     c_map = build_canon_map(df_c.columns)
#     bc_col = pick_first_existing(c_map, ["barcode"])
#     cogs_col = pick_first_existing(c_map, ["cogs rate", "cogs"])
#     if not bc_col or not cogs_col:
#         raise RuntimeError("COGS sheet must have columns: Barcode, COGS Rate")

#     cogs_map = {}
#     for _, row in df_c.iterrows():
#         bc = normalize_ean(row.get(bc_col, ""))
#         if not bc:
#             continue
#         val = pd.to_numeric(row.get(cogs_col, ""), errors="coerce")
#         cogs_map[bc] = float(val) if pd.notna(val) else row.get(cogs_col, "")

#     return rules, cogs_map

# def apply_margin_and_cogs(df_base: pd.DataFrame, margin_rules, cogs_map) -> pd.DataFrame:
#     df = df_base.copy()

#     if "Margin %" not in df.columns:
#         df["Margin %"] = ""
#     if "COGS Rate" not in df.columns:
#         df["COGS Rate"] = ""

#     cust_norm = df["Customer Name"].fillna("").astype(str).map(norm_name)

#     def match_margin(cn: str):
#         if not cn.strip():
#             return ""
#         tokset = set(tokenize(cn))
#         if not tokset:
#             return ""
#         for rule in margin_rules:
#             if rule["tokens"].issubset(tokset):
#                 return rule["margin"]
#         return ""

#     df["Margin %"] = cust_norm.map(match_margin)

#     bcs = df["Barcode"].fillna("").astype(str).map(normalize_ean)
#     df["COGS Rate"] = bcs.map(lambda b: cogs_map.get(b, ""))

#     return df

# # -------------------- 3) Billing calc + reorder --------------------
# def to_percent_fraction(v):
#     if v is None or (isinstance(v, float) and pd.isna(v)):
#         return 0.0
#     if isinstance(v, (int, float)):
#         v = float(v)
#         return v / 100.0 if v > 1 else v
#     s = str(v).strip()
#     if not s:
#         return 0.0
#     if "%" in s:
#         try:
#             return float(s.replace("%", "").strip()) / 100.0
#         except:
#             return 0.0
#     try:
#         n = float(s)
#         return n / 100.0 if n > 1 else n
#     except:
#         return 0.0

# def calc_billing_and_reorder(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()

#     required = ["Customer Name", "Barcode", "Qty.", "MRP Rate", "Margin %", "COGS Rate"]
#     for c in required:
#         if c not in df.columns:
#             raise ValueError(f"Missing required column for billing calc: {c}")

#     qty = pd.to_numeric(df["Qty."], errors="coerce").fillna(0)
#     mrp_rate = pd.to_numeric(df["MRP Rate"], errors="coerce").fillna(0)
#     cogs_rate = pd.to_numeric(df["COGS Rate"], errors="coerce").fillna(0)
#     margin_pct = df["Margin %"].map(to_percent_fraction)

#     gst1_pct = mrp_rate.apply(lambda x: 0.12 if (x > 0 and x <= 1000) else (0.18 if x > 1000 else 0.0))
#     margin_val = mrp_rate * margin_pct
#     gst1 = (mrp_rate * gst1_pct / (1 + gst1_pct)).fillna(0)
#     basic_rate = (mrp_rate - margin_val - gst1).fillna(0)
#     gst2_pct = basic_rate.apply(lambda x: 0.12 if (x > 0 and x <= 1000) else (0.18 if x > 1000 else 0.0))

#     basic_value = basic_rate * qty
#     gst2 = basic_value * gst2_pct
#     bill_value = basic_value + gst2
#     mrp_value = mrp_rate * qty
#     cogs_value = cogs_rate * qty

#     df["GST1%"] = gst1_pct
#     df["Margin"] = margin_val
#     df["GST1"] = gst1
#     df["Basic Rate"] = basic_rate
#     df["GST2%"] = gst2_pct
#     df["Basic Value"] = basic_value
#     df["GST2"] = gst2
#     df["Bill Value"] = bill_value
#     df["MRP Value"] = mrp_value
#     df["COGS Value"] = cogs_value

#     desired = []
#     if "Year" in df.columns:
#         desired.append("Year")
#     if "Month" in df.columns:
#         desired.append("Month")

#     desired += [
#         "Customer Name",
#         "Barcode",
#         "Qty.",
#         "MRP Rate",
#         "Margin %",
#         "GST1%",
#         "GST2%",
#         "MRP Value",
#         "Margin",
#         "GST1",
#         "Basic Rate",
#         "Basic Value",
#         "GST2",
#         "Bill Value",
#         "COGS Rate",
#         "COGS Value",
#     ]

#     extras = [c for c in df.columns if c not in desired]
#     return df[desired + extras]

# # -------------------- 4) Pivot generation --------------------
# def generate_pivot(df_base: pd.DataFrame) -> pd.DataFrame:
#     df = df_base.copy()

#     sum_cols = ["Qty.", "MRP Value", "Basic Value", "GST2", "Bill Value", "COGS Value"]
#     for c in sum_cols:
#         if c in df.columns:
#             df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

#     group_cols = []
#     if "Year" in df.columns:
#         group_cols.append("Year")
#     if "Month" in df.columns:
#         group_cols.append("Month")
#     group_cols.append("Customer Name")

#     pivot = df.groupby(group_cols, as_index=False)[sum_cols].sum()
#     return pivot

# # -------------------- MAIN --------------------
# def main():
#     args = parse_args()

#     input_dir = Path(args.input_dir)
#     output_dir = Path(args.output_dir)
#     output_dir.mkdir(parents=True, exist_ok=True)
#     output_file = output_dir / args.output_name

#     month = (args.month or "").strip().upper()
#     year = (args.year or "").strip()

#     if not input_dir.exists():
#         raise FileNotFoundError(f"Input folder not found: {input_dir}")

#     margin_rules, cogs_map = load_margin_rules_and_cogs_map()
#     print(f"Loaded margin rules: {len(margin_rules)}, cogs rows: {len(cogs_map)}")

#     files = []
#     for ext in ("*.xlsx", "*.xls", "*.xlsb", "*.csv"):
#         files.extend(input_dir.glob(ext))
#     files = sorted(files)

#     if not files:
#         print("No input files found.")
#         return

#     parts = []
#     for p in files:
#         try:
#             print(f"Reading: {p.name}")
#             dfs = read_any_file_all_sheets(p)
#             for df in dfs:
#                 base = ensure_base_cols(df, p)
#                 base = drop_fully_blank_rows(base)
#                 base = add_month_year_cols(base, month, year)
#                 if not base.empty:
#                     parts.append(base)
#         except Exception as e:
#             print(f"Failed: {p.name} -> {e}")

#     if not parts:
#         print("No output generated.")
#         return

#     df_base = pd.concat(parts, ignore_index=True)

#     df_enriched = apply_margin_and_cogs(df_base, margin_rules, cogs_map)
#     df_final = calc_billing_and_reorder(df_enriched)
#     pivot_df = generate_pivot(df_final)

#     with pd.ExcelWriter(output_file, engine="openpyxl") as w:
#         df_final.to_excel(w, sheet_name=DEST_SHEET_NAME, index=False)
#         pivot_df.to_excel(w, sheet_name=PIVOT_SHEET_NAME, index=False)

#     print("Done. Saved:", str(output_file))
#     print(f"{DEST_SHEET_NAME} rows: {len(df_final)}")
#     print(f"{PIVOT_SHEET_NAME} rows: {len(pivot_df)}")

#     load_dotenv()

#     # 1) Push pivot
#     pivot_sheet_id = env_clean(os.getenv("REC_STK_SHEET_ID", ""))
#     pivot_tab_name = env_clean(os.getenv("REC_STK_TAB_NAME", ""))

#     if not pivot_sheet_id or not pivot_tab_name:
#         print("REC_STK_SHEET_ID or REC_STK_TAB_NAME missing in .env. Skipping pivot push.")
#     else:
#         print(
#             f"Upsert/append pivot to Google Sheet tab '{pivot_tab_name}' "
#             f"using Year+Month+Customer Name (Month={month}, Year={year}) ..."
#         )
#         upsert_append_by_month_year_customer(pivot_sheet_id, pivot_tab_name, pivot_df, month, year)
#         print("Pivot Google Sheet update done.")

#     # 2) Push base
#     base_sheet_id = env_clean(os.getenv("REC_STK_BASE_SHEET_ID", ""))
#     base_tab_name = env_clean(os.getenv("REC_STK_BASE_TAB_NAME", ""))

#     if not base_sheet_id or not base_tab_name:
#         print("REC_STK_BASE_SHEET_ID or REC_STK_BASE_TAB_NAME missing in .env. Skipping base push.")
#     else:
#         print(
#             f"Replace/append base data to Google Sheet tab '{base_tab_name}' "
#             f"using full batch replacement by Year+Month (Month={month}, Year={year}) ..."
#         )
#         replace_batch_by_month_year(base_sheet_id, base_tab_name, df_final, month, year)
#         print("Base Google Sheet update done.")

# if __name__ == "__main__":
#     main()


import os
import re
import json
import base64
import math
import argparse
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# ===================== SHEET NAMES =====================
DEST_SHEET_NAME = "Rec_STK_Base"
PIVOT_SHEET_NAME = "Rec_Stk_Final"

# ===================== CLI =====================
def parse_args():
    ap = argparse.ArgumentParser(description="Receivable Stock (rec_stk) processor")
    ap.add_argument("--input_dir", required=True, help="Folder containing input files")
    ap.add_argument("--output_dir", required=True, help="Directory to write output Excel")
    ap.add_argument(
        "--output_name",
        default="REC_STK_MASTER.xlsx",
        help="Output Excel file name (default: REC_STK_MASTER.xlsx)",
    )
    ap.add_argument("--month", default="", help="Month label (e.g. JAN, FEB, ...)")
    ap.add_argument("--year", default="", help="Year (e.g. 2026)")
    return ap.parse_args()

# -------------------- .env helper --------------------
def env_clean(v: str, default: str = "") -> str:
    if v is None:
        return default
    s = str(v).strip()
    if (s.startswith('"') and s.endswith('"')) or (s.startswith("'") and s.endswith("'")):
        s = s[1:-1].strip()
    if s.endswith(";"):
        s = s[:-1].strip()
    return s or default

# -------------------- Common helpers --------------------
def canon(x: str) -> str:
    s = str(x).replace("\u00A0", " ").strip().lower()
    s = re.sub(r"[_\-\.\(\)\[\]\/\\]+", " ", s)
    s = " ".join(s.split())
    return s

def build_canon_map(cols):
    m = {}
    for c in cols:
        m[canon(c)] = c
    return m

def pick_first_existing(col_map, candidates):
    for c in candidates:
        if c in col_map:
            return col_map[c]
    return None

def to_num(s):
    return pd.to_numeric(s, errors="coerce")

def normalize_ean(x) -> str:
    if x is None:
        return ""
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if math.isfinite(x):
            return str(int(round(x)))
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    s = re.sub(r"\.0$", "", s)
    if re.match(r"^\d+(\.\d+)?e\+\d+$", s.lower()):
        try:
            return str(int(round(float(s))))
        except:
            return s
    return s

def infer_customer_from_filename(filename: str) -> str:
    f = canon(filename)
    if "flipkart" in f:
        return "Flipkart"
    if "myntra" in f:
        return "Myntra"
    if "ajio" in f:
        return "Reliance Ajio SOR"
    return ""

def drop_fully_blank_rows(df4: pd.DataFrame) -> pd.DataFrame:
    df4 = df4.copy()
    df4["Customer Name"] = df4["Customer Name"].fillna("").astype(str).str.strip()
    df4["Barcode"] = df4["Barcode"].fillna("").astype(str).str.strip()

    mask_any = (
        df4["Customer Name"].ne("")
        | df4["Barcode"].ne("")
        | df4["Qty."].notna()
        | df4["MRP Rate"].notna()
    )
    return df4[mask_any]

# -------------------- Read files (all sheets) --------------------
def read_any_file_all_sheets(path: Path):
    suf = path.suffix.lower()

    if suf == ".csv":
        try:
            return [pd.read_csv(path)]
        except UnicodeDecodeError:
            return [pd.read_csv(path, encoding="latin1")]

    if suf in (".xlsx", ".xls"):
        xl = pd.ExcelFile(path)
        return [pd.read_excel(path, sheet_name=sn) for sn in xl.sheet_names]

    if suf == ".xlsb":
        xl = pd.ExcelFile(path, engine="pyxlsb")
        return [pd.read_excel(path, sheet_name=sn, engine="pyxlsb") for sn in xl.sheet_names]

    raise ValueError(f"Unsupported file type: {path.name}")

# -------------------- 1) Build base Rec_STK_Base --------------------
def ensure_base_cols(df: pd.DataFrame, file_path: Path) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).replace("\u00A0", " ").strip() for c in df.columns]
    col_map = build_canon_map(df.columns)

    barcode_col = pick_first_existing(col_map, ["eancode", "ean code", "ean", "barcode", "bar code"])
    mrp_col = pick_first_existing(col_map, ["mrp rate", "mrp"])

    fixed = infer_customer_from_filename(file_path.name)

    if fixed == "Flipkart":
        qty_col = pick_first_existing(col_map, ["soh", "soh qty"])
        customer_series = pd.Series(["Flipkart"] * len(df), index=df.index)

    elif fixed == "Myntra":
        qty_col = pick_first_existing(col_map, ["count_", "count"])
        customer_series = pd.Series(["Myntra"] * len(df), index=df.index)

    elif fixed == "Reliance Ajio SOR":
        qty_col = pick_first_existing(col_map, ["count_", "count", "active qty", "active_qty"])
        customer_series = pd.Series(["Reliance Ajio SOR"] * len(df), index=df.index)

    else:
        partner_col = pick_first_existing(col_map, ["partner", "partner name"])
        qty_col = pick_first_existing(col_map, ["closing stock", "closing_stock"])
        customer_series = (
            df[partner_col].fillna("").astype(str).str.strip()
            if partner_col else pd.Series([""] * len(df), index=df.index)
        )

    barcode_series = df[barcode_col].map(normalize_ean) if barcode_col else pd.Series([""] * len(df), index=df.index)
    qty_series = to_num(df[qty_col]) if qty_col else pd.Series([pd.NA] * len(df), index=df.index)
    mrp_series = to_num(df[mrp_col]) if mrp_col else pd.Series([pd.NA] * len(df), index=df.index)

    out = pd.DataFrame({
        "Customer Name": customer_series,
        "Barcode": barcode_series,
        "Qty.": qty_series,
        "MRP Rate": mrp_series,
    })
    return out

def add_month_year_cols(df: pd.DataFrame, month: str, year: str) -> pd.DataFrame:
    df = df.copy()
    month = (month or "").strip().upper()
    year = (year or "").strip()
    if month and year:
        df.insert(0, "Month", month)
        df.insert(0, "Year", year)
    return df

# -------------------- Google Sheets --------------------
def get_gspread_client_from_env(write: bool = False):
    load_dotenv()
    b64 = os.getenv("GOOGLE_SA_JSON_B64")
    if not b64:
        raise RuntimeError("Missing GOOGLE_SA_JSON_B64 in .env (must be one-line base64).")
    sa_json = json.loads(base64.b64decode(b64).decode("utf-8"))
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    if write:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(sa_json, scopes=scopes)
    return gspread.authorize(creds)

def worksheet_to_df(ws) -> pd.DataFrame:
    vals = ws.get_all_values()
    if not vals:
        return pd.DataFrame()
    headers = vals[0]
    rows = vals[1:]
    df = pd.DataFrame(rows, columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _normalize_month_year_for_compare(df: pd.DataFrame):
    if "Year" in df.columns:
        df["Year"] = df["Year"].astype(str).str.strip()
    if "Month" in df.columns:
        df["Month"] = df["Month"].astype(str).str.strip().str.upper()
    return df

def _normalize_customer_name_for_compare(df: pd.DataFrame):
    if "Customer Name" in df.columns:
        df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
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

    existing_df = existing_df[final_cols]
    df_new = df_new[final_cols]
    return existing_df, df_new

def upsert_append_by_month_year_customer(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
    """
    Pivot push logic:
    - match by Year + Month + Customer Name
    - matched existing rows are removed
    - new rows are appended
    - unrelated rows stay untouched
    """
    gc = get_gspread_client_from_env(write=True)
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=1000, cols=max(10, new_df.shape[1] + 2))

    existing_vals = ws.get_all_values()
    if not existing_vals:
        existing_df = pd.DataFrame()
    else:
        headers = existing_vals[0]
        rows = existing_vals[1:]
        existing_df = pd.DataFrame(rows, columns=headers)
        existing_df.columns = [str(c).strip() for c in existing_df.columns]

    df_new = new_df.copy().replace([pd.NA, float("inf"), float("-inf")], "")
    df_new = df_new.where(pd.notnull(df_new), "")

    if existing_df.empty:
        out_df = df_new
    else:
        existing_df, df_new = _align_columns(existing_df, df_new)

        existing_df = _normalize_month_year_for_compare(existing_df)
        df_new = _normalize_month_year_for_compare(df_new)
        existing_df = _normalize_customer_name_for_compare(existing_df)
        df_new = _normalize_customer_name_for_compare(df_new)

        required_cols = {"Year", "Month", "Customer Name"}
        if required_cols.issubset(existing_df.columns) and required_cols.issubset(df_new.columns):
            new_keys = set(
                zip(
                    df_new["Year"].astype(str).str.strip(),
                    df_new["Month"].astype(str).str.upper(),
                    df_new["Customer Name"].astype(str).str.strip(),
                )
            )

            existing_keys = list(
                zip(
                    existing_df["Year"].astype(str).str.strip(),
                    existing_df["Month"].astype(str).str.upper(),
                    existing_df["Customer Name"].astype(str).str.strip(),
                )
            )

            mask_same = [key in new_keys for key in existing_keys]
            existing_df = existing_df.loc[[not matched for matched in mask_same]].copy()

        out_df = pd.concat([existing_df, df_new], ignore_index=True)

    if {"Year", "Month", "Customer Name"}.issubset(out_df.columns):
        out_df["Year"] = out_df["Year"].astype(str).str.strip()
        out_df["Month"] = out_df["Month"].astype(str).str.strip().str.upper()
        out_df["Customer Name"] = out_df["Customer Name"].astype(str).str.strip()
        out_df = out_df.drop_duplicates(subset=["Year", "Month", "Customer Name"], keep="last")

    out_df = out_df.where(pd.notnull(out_df), "")
    values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")

def replace_batch_by_month_year(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
    """
    Base push logic:
    - remove all existing rows for the same Year + Month
    - append the new batch for that Year + Month
    - other months remain untouched
    """
    month = (month or "").strip().upper()
    year = (year or "").strip()

    gc = get_gspread_client_from_env(write=True)
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=1000, cols=max(10, new_df.shape[1] + 2))

    existing_vals = ws.get_all_values()
    if not existing_vals:
        existing_df = pd.DataFrame()
    else:
        headers = existing_vals[0]
        rows = existing_vals[1:]
        existing_df = pd.DataFrame(rows, columns=headers)
        existing_df.columns = [str(c).strip() for c in existing_df.columns]

    df_new = new_df.copy().replace([pd.NA, float("inf"), float("-inf")], "")
    df_new = df_new.where(pd.notnull(df_new), "")

    if existing_df.empty:
        out_df = df_new
    else:
        existing_df, df_new = _align_columns(existing_df, df_new)

        existing_df = _normalize_month_year_for_compare(existing_df)
        df_new = _normalize_month_year_for_compare(df_new)
        existing_df = _normalize_customer_name_for_compare(existing_df)
        df_new = _normalize_customer_name_for_compare(df_new)

        if month and year and {"Year", "Month"}.issubset(existing_df.columns):
            mask_same_batch = (
                (existing_df["Year"].astype(str).str.strip() == year) &
                (existing_df["Month"].astype(str).str.upper() == month)
            )
            existing_df = existing_df.loc[~mask_same_batch].copy()

        out_df = pd.concat([existing_df, df_new], ignore_index=True)

    out_df = out_df.where(pd.notnull(out_df), "")
    values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")

# -------------------- 2) Margin rules + COGS map --------------------
STOP_WORDS = {
    "PVT", "PRIVATE", "LTD", "LIMITED", "LLP", "INDIA",
    "INTERNATIONAL", "GLOBAL", "VENTURE", "PVT.", "LTD.",
    "THE", "AND"
}

GENERIC_AMBIGUOUS_TOKENS = {
    "RELIANCE", "RETAIL", "STORE", "SHOP", "ONLINE", "SOR"
}

def norm_name(s: str) -> str:
    s = str(s or "").upper().strip()
    s = s.replace("&", " AND ")
    s = re.sub(r"[_\-/\\]+", " ", s)
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def singular_word(w: str) -> str:
    if len(w) > 3 and w.endswith("S"):
        return w[:-1]
    return w

def tokenize(normed: str):
    toks = [t.strip() for t in str(normed).split(" ") if t.strip()]
    toks = [singular_word(t) for t in toks]
    toks = [t for t in toks if t and t not in STOP_WORDS]
    return toks

def to_percent_number(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return ""
        v = float(v)
        return v / 100.0 if v > 1 else v
    s = str(v).strip()
    if not s:
        return ""
    if "%" in s:
        try:
            return float(s.replace("%", "").strip()) / 100.0
        except:
            return ""
    try:
        n = float(s)
        return n / 100.0 if n > 1 else n
    except:
        return ""

def generate_aliases(master_name_raw: str):
    """
    Build safe aliases for a master customer.
    Returns normalized alias strings.
    """
    raw = str(master_name_raw or "").strip()
    if not raw:
        return set()

    aliases = set()

    full_norm = norm_name(raw)
    if full_norm:
        aliases.add(full_norm)

    split_parts = [p.strip() for p in re.split(r"[_\-/\\]+", raw) if p.strip()]
    for p in split_parts:
        pn = norm_name(p)
        if pn:
            aliases.add(pn)

    full_upper = full_norm

    if full_upper == "RELIANCE CENTRO":
        aliases.update({
            "CENTRO",
            "RELIANCE CENTRO",
        })

    elif full_upper == "RELIANCE AJIO SOR":
        aliases.update({
            "AJIO",
            "AJIO SOR",
            "RELIANCE AJIO",
            "RELIANCE AJIO SOR",
        })

    elif full_upper == "LEAYAN ZUUP":
        aliases.update({
            "LEAYAN",
            "ZUUP",
            "LEAYAN ZUUP",
        })

    return {a for a in aliases if a}

def build_margin_index(df_m: pd.DataFrame, cust_col: str, margin_col: str):
    """
    Creates:
    1) exact_map -> exact normalized alias to margin
    2) token_rules -> strong token-based rules
    """
    exact_map = {}
    token_rules = []

    for _, row in df_m.iterrows():
        master_name_raw = str(row.get(cust_col, "")).strip()
        if not master_name_raw:
            continue

        margin_val = to_percent_number(row.get(margin_col, ""))
        if margin_val == "":
            continue

        aliases = generate_aliases(master_name_raw)

        for alias in aliases:
            exact_map[alias] = float(margin_val)

            toks = tokenize(alias)
            if toks:
                token_rules.append({
                    "alias": alias,
                    "tokens": set(toks),
                    "margin": float(margin_val),
                    "score": len(set(toks)) * 1000 + len(alias),
                })

    token_rules.sort(key=lambda r: r["score"], reverse=True)
    return exact_map, token_rules

def load_margin_rules_and_cogs_map():
    load_dotenv()
    source_spreadsheet_id = env_clean(os.getenv("SOURCE_SPREADSHEET_ID1", ""))
    margin_sheet_name = env_clean(os.getenv("MARGIN_SHEET_NAME1", "Margin"))
    cogs_sheet_name = env_clean(os.getenv("COGS_SHEET_NAME1", "COGS"))

    if not source_spreadsheet_id:
        raise RuntimeError("Missing SOURCE_SPREADSHEET_ID1 in .env")

    gc = get_gspread_client_from_env()
    sh = gc.open_by_key(source_spreadsheet_id)

    ws_m = sh.worksheet(margin_sheet_name)
    df_m = worksheet_to_df(ws_m)
    if df_m.empty:
        raise RuntimeError("Margin sheet empty")

    m_map = build_canon_map(df_m.columns)
    cust_col = pick_first_existing(m_map, ["customer name", "customer"])
    margin_col = pick_first_existing(m_map, ["margin"])
    if not cust_col or not margin_col:
        raise RuntimeError("Margin sheet must have columns: Customer Name, Margin")

    margin_exact_map, margin_token_rules = build_margin_index(df_m, cust_col, margin_col)

    ws_c = sh.worksheet(cogs_sheet_name)
    df_c = worksheet_to_df(ws_c)
    if df_c.empty:
        raise RuntimeError("COGS sheet empty")

    c_map = build_canon_map(df_c.columns)
    bc_col = pick_first_existing(c_map, ["barcode"])
    cogs_col = pick_first_existing(c_map, ["cogs rate", "cogs"])
    if not bc_col or not cogs_col:
        raise RuntimeError("COGS sheet must have columns: Barcode, COGS Rate")

    cogs_map = {}
    for _, row in df_c.iterrows():
        bc = normalize_ean(row.get(bc_col, ""))
        if not bc:
            continue
        val = pd.to_numeric(row.get(cogs_col, ""), errors="coerce")
        cogs_map[bc] = float(val) if pd.notna(val) else row.get(cogs_col, "")

    return margin_exact_map, margin_token_rules, cogs_map

def resolve_margin(customer_name: str, margin_exact_map, margin_token_rules):
    """
    Matching priority:
    1. exact normalized match
    2. exact alias match
    3. safe token subset match
    4. ambiguous -> blank
    """
    cn_norm = norm_name(customer_name)
    if not cn_norm:
        return ""

    if cn_norm in margin_exact_map:
        return margin_exact_map[cn_norm]

    tokset = set(tokenize(cn_norm))
    if not tokset:
        return ""

    if len(tokset) == 1 and list(tokset)[0] in GENERIC_AMBIGUOUS_TOKENS:
        return ""

    matches = []
    for rule in margin_token_rules:
        rule_tokens = rule["tokens"]

        if tokset == rule_tokens:
            matches.append(rule)
            continue

        if rule_tokens.issubset(tokset):
            matches.append(rule)
            continue

        if tokset.issubset(rule_tokens):
            if len(tokset) == 1:
                only_tok = next(iter(tokset))
                if only_tok in GENERIC_AMBIGUOUS_TOKENS:
                    continue
            matches.append(rule)

    if not matches:
        return ""

    matches.sort(key=lambda x: x["score"], reverse=True)

    if len(matches) > 1:
        top = matches[0]
        second = matches[1]
        if top["score"] == second["score"] and top["margin"] != second["margin"]:
            return ""

    return matches[0]["margin"]

def apply_margin_and_cogs(df_base: pd.DataFrame, margin_exact_map, margin_token_rules, cogs_map) -> pd.DataFrame:
    df = df_base.copy()

    if "Margin %" not in df.columns:
        df["Margin %"] = ""
    if "COGS Rate" not in df.columns:
        df["COGS Rate"] = ""

    df["Margin %"] = df["Customer Name"].fillna("").astype(str).map(
        lambda x: resolve_margin(x, margin_exact_map, margin_token_rules)
    )

    bcs = df["Barcode"].fillna("").astype(str).map(normalize_ean)
    df["COGS Rate"] = bcs.map(lambda b: cogs_map.get(b, ""))

    return df

# -------------------- 3) Billing calc + reorder --------------------
def to_percent_fraction(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        v = float(v)
        return v / 100.0 if v > 1 else v
    s = str(v).strip()
    if not s:
        return 0.0
    if "%" in s:
        try:
            return float(s.replace("%", "").strip()) / 100.0
        except:
            return 0.0
    try:
        n = float(s)
        return n / 100.0 if n > 1 else n
    except:
        return 0.0

def calc_billing_and_reorder(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    required = ["Customer Name", "Barcode", "Qty.", "MRP Rate", "Margin %", "COGS Rate"]
    for c in required:
        if c not in df.columns:
            raise ValueError(f"Missing required column for billing calc: {c}")

    qty = pd.to_numeric(df["Qty."], errors="coerce").fillna(0)
    mrp_rate = pd.to_numeric(df["MRP Rate"], errors="coerce").fillna(0)
    cogs_rate = pd.to_numeric(df["COGS Rate"], errors="coerce").fillna(0)
    margin_pct = df["Margin %"].map(to_percent_fraction)

    gst1_pct = mrp_rate.apply(lambda x: 0.12 if (x > 0 and x <= 1000) else (0.18 if x > 1000 else 0.0))
    margin_val = mrp_rate * margin_pct
    gst1 = (mrp_rate * gst1_pct / (1 + gst1_pct)).fillna(0)
    basic_rate = (mrp_rate - margin_val - gst1).fillna(0)
    gst2_pct = basic_rate.apply(lambda x: 0.12 if (x > 0 and x <= 1000) else (0.18 if x > 1000 else 0.0))

    basic_value = basic_rate * qty
    gst2 = basic_value * gst2_pct
    bill_value = basic_value + gst2
    mrp_value = mrp_rate * qty
    cogs_value = cogs_rate * qty

    df["GST1%"] = gst1_pct
    df["Margin"] = margin_val
    df["GST1"] = gst1
    df["Basic Rate"] = basic_rate
    df["GST2%"] = gst2_pct
    df["Basic Value"] = basic_value
    df["GST2"] = gst2
    df["Bill Value"] = bill_value
    df["MRP Value"] = mrp_value
    df["COGS Value"] = cogs_value

    desired = []
    if "Year" in df.columns:
        desired.append("Year")
    if "Month" in df.columns:
        desired.append("Month")

    desired += [
        "Customer Name",
        "Barcode",
        "Qty.",
        "MRP Rate",
        "Margin %",
        "GST1%",
        "GST2%",
        "MRP Value",
        "Margin",
        "GST1",
        "Basic Rate",
        "Basic Value",
        "GST2",
        "Bill Value",
        "COGS Rate",
        "COGS Value",
    ]

    extras = [c for c in df.columns if c not in desired]
    return df[desired + extras]

# -------------------- 4) Pivot generation --------------------
def generate_pivot(df_base: pd.DataFrame) -> pd.DataFrame:
    df = df_base.copy()

    sum_cols = ["Qty.", "MRP Value", "Basic Value", "GST2", "Bill Value", "COGS Value"]
    for c in sum_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    group_cols = []
    if "Year" in df.columns:
        group_cols.append("Year")
    if "Month" in df.columns:
        group_cols.append("Month")
    group_cols.append("Customer Name")

    pivot = df.groupby(group_cols, as_index=False)[sum_cols].sum()
    return pivot

# -------------------- MAIN --------------------
def main():
    args = parse_args()

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / args.output_name

    month = (args.month or "").strip().upper()
    year = (args.year or "").strip()

    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    margin_exact_map, margin_token_rules, cogs_map = load_margin_rules_and_cogs_map()
    print(
        f"Loaded margin aliases: {len(margin_exact_map)}, "
        f"margin token rules: {len(margin_token_rules)}, "
        f"cogs rows: {len(cogs_map)}"
    )

    files = []
    for ext in ("*.xlsx", "*.xls", "*.xlsb", "*.csv"):
        files.extend(input_dir.glob(ext))
    files = sorted(files)

    if not files:
        print("No input files found.")
        return

    parts = []
    for p in files:
        try:
            print(f"Reading: {p.name}")
            dfs = read_any_file_all_sheets(p)
            for df in dfs:
                base = ensure_base_cols(df, p)
                base = drop_fully_blank_rows(base)
                base = add_month_year_cols(base, month, year)
                if not base.empty:
                    parts.append(base)
        except Exception as e:
            print(f"Failed: {p.name} -> {e}")

    if not parts:
        print("No output generated.")
        return

    df_base = pd.concat(parts, ignore_index=True)

    df_enriched = apply_margin_and_cogs(df_base, margin_exact_map, margin_token_rules, cogs_map)
    df_final = calc_billing_and_reorder(df_enriched)
    pivot_df = generate_pivot(df_final)

    with pd.ExcelWriter(output_file, engine="openpyxl") as w:
        df_final.to_excel(w, sheet_name=DEST_SHEET_NAME, index=False)
        pivot_df.to_excel(w, sheet_name=PIVOT_SHEET_NAME, index=False)

    print("Done. Saved:", str(output_file))
    print(f"{DEST_SHEET_NAME} rows: {len(df_final)}")
    print(f"{PIVOT_SHEET_NAME} rows: {len(pivot_df)}")

    load_dotenv()

    # 1) Push pivot
    pivot_sheet_id = env_clean(os.getenv("REC_STK_SHEET_ID", ""))
    pivot_tab_name = env_clean(os.getenv("REC_STK_TAB_NAME", ""))

    if not pivot_sheet_id or not pivot_tab_name:
        print("REC_STK_SHEET_ID or REC_STK_TAB_NAME missing in .env. Skipping pivot push.")
    else:
        print(
            f"Upsert/append pivot to Google Sheet tab '{pivot_tab_name}' "
            f"using Year+Month+Customer Name (Month={month}, Year={year}) ..."
        )
        upsert_append_by_month_year_customer(pivot_sheet_id, pivot_tab_name, pivot_df, month, year)
        print("Pivot Google Sheet update done.")

    # 2) Push base
    base_sheet_id = env_clean(os.getenv("REC_STK_BASE_SHEET_ID", ""))
    base_tab_name = env_clean(os.getenv("REC_STK_BASE_TAB_NAME", ""))

    if not base_sheet_id or not base_tab_name:
        print("REC_STK_BASE_SHEET_ID or REC_STK_BASE_TAB_NAME missing in .env. Skipping base push.")
    else:
        print(
            f"Replace/append base data to Google Sheet tab '{base_tab_name}' "
            f"using full batch replacement by Year+Month (Month={month}, Year={year}) ..."
        )
        replace_batch_by_month_year(base_sheet_id, base_tab_name, df_final, month, year)
        print("Base Google Sheet update done.")

if __name__ == "__main__":
    main()