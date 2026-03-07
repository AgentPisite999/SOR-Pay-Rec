

# import os
# import re
# import json
# import base64
# import argparse
# import sys
# from pathlib import Path

# import pandas as pd
# from dotenv import load_dotenv
# import gspread
# from google.oauth2.service_account import Credentials

# # ===================== FIX: FORCE UTF-8 PRINTING (prevents ✅/❌ crash on Windows) =====================
# try:
#     sys.stdout.reconfigure(encoding="utf-8")
#     sys.stderr.reconfigure(encoding="utf-8")
# except Exception:
#     pass

# # ===================== DEFAULTS (used only if CLI not provided) =====================
# DEFAULT_INPUT_DIR = Path(r"C:\path\to\input")
# DEFAULT_OUTPUT_NAME = "MASTER_OUTPUT.xlsx"

# RAW_SHEET_NAME = "Receivable_Raw"
# PIVOT_SHEET_NAME = "Final_Rec"
# EXCLUDE_SHEETS = {RAW_SHEET_NAME, PIVOT_SHEET_NAME, "Remark"}

# REQUIRED_HEADERS = [
#     "Partner Name",
#     "Qty (Sales)",
#     "Gross Sales (Sales)",
#     "RSP (Sales)",
#     "Rebate (Sales)",
#     "Discount_Approved (Sales)",
#     "Discount Not Approved (Sales)",
#     "Total Discount (Sales)",
#     "Net Sales (Sales)",
#     "Basic Value (Billing Cost )",
#     "GST Reimbursement (Billing Cost )",
#     "Bill Value (Billing Cost )",
#     "Basic Value (Payment )",
#     "GST Reimbursement (Payment )",
#     "Bill Value (Payment )",
#     "Discount CN",
#     "Margin Value (Billing Rec)",
#     "GST Payable (Billing Rec)",
#     "Margin Value (Payment Rec)",
#     "GST Payable (Payment Rec)",
# ]

# # ---------- Partner inference from filename ----------
# PARTNER_KEYWORDS = [
#     ("myntra", "Myntra jabong"),
#     ("jabong", "Myntra jabong"),
#     ("flipkart", "Flipkart"),
#     ("v - retail", "V - Retail"),
#     ("v-retail", "V - Retail"),
#     ("kora", "Kora"),
#     ("zuup", "ZUUP"),
#     ("nirankar", "NIRANKAR"),
#     ("shoppers stop", "SHOPPERS STOP"),
#     ("shoppersstop", "SHOPPERS STOP"),
#     ("relience centro", "Relience Centro"),
#     ("reliance centro", "Relience Centro"),
#     ("reliance retail ajio sor", "Reliance Retail Ajio SOR"),
#     ("ajio sor", "Reliance Retail Ajio SOR"),
#     ("ajio", "Reliance Retail Ajio SOR"),
#     (" ls ", "LS"),  # keep last
# ]

# # -------------------- CLI --------------------
# def parse_args():
#     ap = argparse.ArgumentParser(description="Receivable Trade Discount processor (multi-file)")
#     ap.add_argument("--input_dir", required=True, help="Folder containing input Excel files")
#     ap.add_argument("--output_dir", required=True, help="Folder to write output file")
#     ap.add_argument("--output_name", default=DEFAULT_OUTPUT_NAME, help="Output Excel file name")
#     ap.add_argument("--month", default="", help="Month label (e.g. JAN, FEB, ...)")
#     ap.add_argument("--year", default="", help="Year (e.g. 2026)")
#     return ap.parse_args()


# # -------------------- Common helpers --------------------
# def canon(name: str) -> str:
#     return " ".join(str(name).strip().lower().split())


# def build_canon_map(columns):
#     """
#     Canonical -> actual. If duplicates exist, LAST one wins (your requirement: use 2nd block).
#     """
#     m = {}
#     for col in columns:
#         m[canon(col)] = col
#     return m


# def to_num(s: pd.Series) -> pd.Series:
#     """
#     Robust numeric conversion:
#     - removes commas: '2,890' -> '2890'
#     - handles NBSP
#     - trims spaces
#     - converts blanks to 0
#     """
#     if s is None:
#         return pd.Series(dtype="float64")
#     if not isinstance(s, pd.Series):
#         s = pd.Series(s)

#     cleaned = (
#         s.astype(str)
#         .str.replace("\u00A0", " ", regex=False)
#         .str.replace(",", "", regex=False)
#         .str.strip()
#     )
#     cleaned = cleaned.replace({"": None, "nan": None, "None": None, "NaN": None})
#     return pd.to_numeric(cleaned, errors="coerce").fillna(0)


# def normalize_partner(name) -> str:
#     return re.sub(r"[\s\-]+", "", str(name).lower()).strip()


# def sanitize_excel_sheet_name(name: str) -> str:
#     name = str(name).strip() or "Sheet"
#     name = re.sub(r"[:\\/?*\[\]]", " ", name)
#     name = re.sub(r"\s+", " ", name).strip() or "Sheet"
#     return name[:31]


# def make_unique_sheet_name(existing: set, desired: str) -> str:
#     if desired not in existing:
#         existing.add(desired)
#         return desired
#     i = 2
#     while True:
#         suffix = f"_{i}"
#         base = desired[: (31 - len(suffix))]
#         candidate = f"{base}{suffix}"
#         if candidate not in existing:
#             existing.add(candidate)
#             return candidate
#         i += 1


# def as_decimal_percent(x):
#     s = str(x).strip().replace("%", "")
#     try:
#         v = float(s)
#     except:
#         return 0.0
#     return v / 100.0 if v > 1.5 else v


# def get_gst_rates_from_threshold(threshold_value: float):
#     if threshold_value == 2625 or threshold_value == 2500:
#         return {"low": 0.05, "high": 0.18}
#     return {"low": 0.12, "high": 0.18}


# def normalize_ean(x) -> str:
#     """
#     Normalize EAN/GTIN for matching COGS master:
#     - strips spaces / NBSP
#     - removes trailing .0
#     - converts scientific notation (8.90913E+12) -> full integer string
#     """
#     if x is None:
#         return ""
#     s = str(x).replace("\u00A0", " ").strip()
#     s = re.sub(r"\.0$", "", s)
#     if re.match(r"^\d+(\.\d+)?e\+\d+$", s.lower()):
#         try:
#             s = str(int(float(s)))
#         except:
#             pass
#     return s


# def infer_partner_from_filename(filename: str) -> str:
#     f = str(filename).lower()
#     f = f.replace("_", " ").replace("-", " ")
#     f = " ".join(f.split())
#     f_pad = f" {f} "
#     for key, partner in PARTNER_KEYWORDS:
#         if key.strip() == "ls":
#             if " ls " in f_pad:
#                 return partner
#         else:
#             if key in f:
#                 return partner
#     return ""


# def ensure_partner_name(df: pd.DataFrame, source_filename: str) -> pd.DataFrame:
#     df = df.copy()
#     inferred = infer_partner_from_filename(source_filename)
#     if not inferred:
#         return df

#     col_map = build_canon_map(df.columns)

#     if "partner name" not in col_map:
#         df["Partner Name"] = inferred
#         return df

#     partner_col = col_map["partner name"]
#     s = (
#         df[partner_col]
#         .astype(str)
#         .fillna("")
#         .str.replace("\u00A0", " ", regex=False)
#         .str.strip()
#     )
#     mask_blank = (s == "") | (s.str.lower() == "nan")
#     df.loc[mask_blank, partner_col] = inferred
#     return df


# def env_clean(v: str, default: str = "") -> str:
#     """
#     Make .env tolerant: strips surrounding quotes and trailing semicolons.
#     So it works even if you wrote: NAME="Value";
#     """
#     if v is None:
#         return default
#     s = str(v).strip()
#     s = s.strip('"').strip("'").strip()
#     if s.endswith(";"):
#         s = s[:-1].strip()
#     return s or default


# # -------------------- Month/Year helper --------------------
# def add_month_year_cols(df: pd.DataFrame, month: str, year: str) -> pd.DataFrame:
#     """
#     Inserts Year + Month as first two columns (constant for this run),
#     only if BOTH provided.
#     """
#     df = df.copy()
#     month = (month or "").strip().upper()
#     year = (year or "").strip()
#     if month and year:
#         if "Year" in df.columns:
#             df["Year"] = year
#         else:
#             df.insert(0, "Year", year)

#         if "Month" in df.columns:
#             df["Month"] = month
#         else:
#             df.insert(1, "Month", month)
#     return df


# # -------------------- Google Sheets (Service Account) --------------------
# def get_gspread_client_from_env(write: bool = False):
#     load_dotenv(override=True)

#     b64 = os.getenv("GOOGLE_SA_JSON_B64")
#     if not b64:
#         raise RuntimeError("Missing GOOGLE_SA_JSON_B64 in .env (must be ONE LINE, no line breaks).")

#     sa_json = json.loads(base64.b64decode(b64).decode("utf-8"))
#     scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
#     if write:
#         scopes = ["https://www.googleapis.com/auth/spreadsheets"]

#     creds = Credentials.from_service_account_info(sa_json, scopes=scopes)
#     return gspread.authorize(creds)


# def worksheet_to_df(ws) -> pd.DataFrame:
#     values = ws.get_all_values()
#     if not values:
#         return pd.DataFrame()
#     headers = values[0]
#     rows = values[1:]
#     df = pd.DataFrame(rows, columns=headers)
#     df.columns = [str(c).strip() for c in df.columns]
#     return df


# def _normalize_month_year_for_compare(df: pd.DataFrame) -> pd.DataFrame:
#     if "Year" in df.columns:
#         df["Year"] = df["Year"].astype(str).str.strip()
#     if "Month" in df.columns:
#         df["Month"] = df["Month"].astype(str).str.strip().str.upper()
#     return df


# def upsert_append_by_month_year(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
#     """
#     Rules:
#     - If month+year provided & sheet has Month+Year columns:
#         remove existing rows matching that month+year, then append new rows
#     - Else:
#         append new rows
#     Does NOT delete other months.
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
#         for c in existing_df.columns:
#             if c not in df_new.columns:
#                 df_new[c] = ""
#         for c in df_new.columns:
#             if c not in existing_df.columns:
#                 existing_df[c] = ""

#         existing_cols = list(existing_df.columns)
#         extra_cols = [c for c in df_new.columns if c not in existing_cols]
#         final_cols = existing_cols + extra_cols

#         existing_df = existing_df[final_cols]
#         df_new = df_new[final_cols]

#         existing_df = _normalize_month_year_for_compare(existing_df)
#         df_new = _normalize_month_year_for_compare(df_new)

#         if month and year and ("Month" in existing_df.columns) and ("Year" in existing_df.columns):
#             mask_same = (existing_df["Month"] == month) & (existing_df["Year"] == year)
#             existing_df = existing_df.loc[~mask_same].copy()

#         out_df = pd.concat([existing_df, df_new], ignore_index=True)

#     out_df = out_df.where(pd.notnull(out_df), "")
#     values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

#     ws.clear()
#     ws.update(values, value_input_option="USER_ENTERED")


# def load_google_maps_and_cogs():
#     load_dotenv(override=True)

#     source_spreadsheet_id = env_clean(os.getenv("MARGIN_SPREADSHEET_ID2", ""))
#     margin_sheet_name = env_clean(os.getenv("MARGIN_SHEET_NAME2", "Margine"))
#     gst_threshold_sheet_name = env_clean(os.getenv("GST_THRESHOLD_SHEET_NAME2", "GST Threshold"))
#     cogs_sheet_name = env_clean(os.getenv("COGS_SHEET_NAME2", "COGS Master"))

#     if not source_spreadsheet_id:
#         raise RuntimeError("Missing MARGIN_SPREADSHEET_ID2 in .env")

#     gc = get_gspread_client_from_env(write=False)
#     sh = gc.open_by_key(source_spreadsheet_id)

#     # ---- Margine ----
#     ws_m = sh.worksheet(margin_sheet_name)
#     df_m = worksheet_to_df(ws_m)
#     if df_m.empty:
#         raise RuntimeError(f"No data found in sheet: {margin_sheet_name}")

#     m_map = build_canon_map(df_m.columns)
#     need_cols = ["partner name", "billing", "payment"]
#     if any(c not in m_map for c in need_cols):
#         raise RuntimeError("Margine sheet must have columns: Partner Name, Billing, Payment")

#     col_partner = m_map["partner name"]
#     col_billing = m_map["billing"]
#     col_payment = m_map["payment"]

#     margin_billing_map = {}
#     margin_payment_map = {}

#     for _, row in df_m.iterrows():
#         p = row.get(col_partner, "")
#         if not str(p).strip():
#             continue
#         key = normalize_partner(p)

#         b = row.get(col_billing, "")
#         py = row.get(col_payment, "")

#         if str(b).strip():
#             margin_billing_map[key] = as_decimal_percent(b)
#         if str(py).strip():
#             margin_payment_map[key] = as_decimal_percent(py)

#     # ---- GST Threshold ----
#     ws_g = sh.worksheet(gst_threshold_sheet_name)
#     df_g = worksheet_to_df(ws_g)
#     if df_g.empty:
#         raise RuntimeError(f"No data found in sheet: {gst_threshold_sheet_name}")

#     g_map = build_canon_map(df_g.columns)

#     need_b = ["partner name", "gst payable (billing cost ) %", "gst reimbursement (billing cost ) %"]
#     need_p = ["partner name", "gst payable (payment ) %", "gst reimbursement (payment ) %"]

#     if any(c not in g_map for c in need_b):
#         raise RuntimeError("GST Threshold missing Billing Cost columns")
#     if any(c not in g_map for c in need_p):
#         raise RuntimeError("GST Threshold missing Payment columns")

#     pcol = g_map["partner name"]
#     b_pay = g_map["gst payable (billing cost ) %"]
#     b_rei = g_map["gst reimbursement (billing cost ) %"]
#     p_pay = g_map["gst payable (payment ) %"]
#     p_rei = g_map["gst reimbursement (payment ) %"]

#     gst_billing_threshold_map = {}
#     gst_payment_threshold_map = {}

#     for _, row in df_g.iterrows():
#         p = row.get(pcol, "")
#         if not str(p).strip():
#             continue
#         key = normalize_partner(p)

#         pay_b = pd.to_numeric(str(row.get(b_pay, "")).replace(",", ""), errors="coerce")
#         rei_b = pd.to_numeric(str(row.get(b_rei, "")).replace(",", ""), errors="coerce")
#         gst_billing_threshold_map[key] = {
#             "payBillingThresh": float(pay_b) if pd.notna(pay_b) else 0.0,
#             "reimbBillingThresh": float(rei_b) if pd.notna(rei_b) else 0.0,
#         }

#         pay_p = pd.to_numeric(str(row.get(p_pay, "")).replace(",", ""), errors="coerce")
#         rei_p = pd.to_numeric(str(row.get(p_rei, "")).replace(",", ""), errors="coerce")
#         gst_payment_threshold_map[key] = {
#             "payPayThresh": float(pay_p) if pd.notna(pay_p) else 0.0,
#             "reimbPayThresh": float(rei_p) if pd.notna(rei_p) else 0.0,
#         }

#     # ---- COGS Master ----
#     ws_c = sh.worksheet(cogs_sheet_name)
#     df_c = worksheet_to_df(ws_c)
#     if df_c.empty:
#         raise RuntimeError(f"No data found in sheet: {cogs_sheet_name}")

#     c_map = build_canon_map(df_c.columns)
#     if "eancode" not in c_map or "rate" not in c_map:
#         raise RuntimeError("COGS Master must have columns: EANCODE, Rate")

#     e_col = c_map["eancode"]
#     r_col = c_map["rate"]

#     cogs_map = {}
#     for _, row in df_c.iterrows():
#         ean = normalize_ean(row.get(e_col, ""))
#         if not ean:
#             continue
#         rate = pd.to_numeric(str(row.get(r_col, "")).replace(",", ""), errors="coerce")
#         cogs_map[ean] = float(rate) if pd.notna(rate) else 0.0

#     return margin_billing_map, margin_payment_map, gst_billing_threshold_map, gst_payment_threshold_map, cogs_map


# # -------------------- 1) SALES --------------------
# SALES_REQUIRED = ["qty", "gross mrp", "rsp", "approved"]

# def process_sales(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     col_map = build_canon_map(df.columns)
#     if any(req not in col_map for req in SALES_REQUIRED):
#         return df

#     qty = to_num(df[col_map["qty"]])
#     gross_mrp = to_num(df[col_map["gross mrp"]])
#     rsp = to_num(df[col_map["rsp"]])
#     approved = to_num(df[col_map["approved"]])

#     qty_sales = qty
#     gross_sales = gross_mrp
#     rsp_sales = rsp.where(qty >= 0, rsp * -1)

#     calculated_rebate = gross_sales - rsp_sales

#     rebate = pd.Series(0.0, index=df.index)
#     disc_approved = pd.Series(0.0, index=df.index)

#     mask = approved < calculated_rebate
#     rebate[mask] = 0
#     disc_approved[mask] = approved[mask]

#     rebate[~mask] = calculated_rebate[~mask]
#     disc_approved[~mask] = approved[~mask] - rebate[~mask]

#     total_discount = rebate + disc_approved
#     net_sales = gross_sales - total_discount

#     df["Qty (Sales)"] = qty_sales.round(0)
#     df["Gross Sales (Sales)"] = gross_sales.round(2)
#     df["RSP (Sales)"] = rsp_sales.round(2)
#     df["Rebate (Sales)"] = rebate.round(2)
#     df["Discount_Approved (Sales)"] = disc_approved.round(2)
#     df["Discount Not Approved (Sales)"] = ""
#     df["Total Discount (Sales)"] = total_discount.round(2)
#     df["Net Sales (Sales)"] = net_sales.round(2)
#     return df


# # -------------------- 2) BILLING COST --------------------
# BILL_REQUIRED = ["gross sales (sales)", "qty (sales)", "partner name"]

# def process_billing_cost(df: pd.DataFrame, margin_billing_map: dict, gst_billing_threshold_map: dict) -> pd.DataFrame:
#     df = df.copy()
#     col_map = build_canon_map(df.columns)
#     if any(req not in col_map for req in BILL_REQUIRED):
#         return df

#     gross_sales = to_num(df[col_map["gross sales (sales)"]])
#     qty_sales = to_num(df[col_map["qty (sales)"]])
#     partner_series = df[col_map["partner name"]].fillna("")
#     partner_key = partner_series.map(normalize_partner)

#     margin_perc = partner_key.map(lambda k: margin_billing_map.get(k, 0.0)).astype(float)

#     mrp_rate = pd.Series(0.0, index=df.index)
#     nz = qty_sales != 0
#     mrp_rate[nz] = gross_sales[nz] / qty_sales[nz]

#     thresh_obj = partner_key.map(lambda k: gst_billing_threshold_map.get(k, {}))
#     pay_thresh = thresh_obj.map(lambda d: d.get("payBillingThresh", 0.0)).astype(float)
#     reimb_thresh = thresh_obj.map(lambda d: d.get("reimbBillingThresh", 0.0)).astype(float)

#     pay_thresh = pay_thresh.where(pay_thresh != 0, 1000.0)
#     reimb_thresh = reimb_thresh.where(reimb_thresh != 0, pay_thresh)

#     pay_low = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
#     pay_high = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
#     gst_pay_perc = pd.Series(
#         [low if rate <= thr else high for rate, thr, low, high in zip(mrp_rate, pay_thresh, pay_low, pay_high)],
#         index=df.index,
#     )

#     margin_value = mrp_rate * margin_perc
#     gst_pay_value = (mrp_rate * gst_pay_perc / (gst_pay_perc + 1)).fillna(0)
#     basic_rate = mrp_rate - margin_value - gst_pay_value

#     reimb_low = reimb_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
#     reimb_high = reimb_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
#     gst_reimb_perc = pd.Series(
#         [low if b <= thr else high for b, thr, low, high in zip(basic_rate, reimb_thresh, reimb_low, reimb_high)],
#         index=df.index,
#     )

#     basic_val = basic_rate * qty_sales
#     gst_reimb_val = basic_val * gst_reimb_perc
#     bill_val = basic_val + gst_reimb_val

#     df["Margin (Billing Cost ) %"] = margin_perc.round(4)
#     df["MRP Rate (Billing Cost ) %"] = mrp_rate.round(2)
#     df["GST Payable (Billing Cost ) %"] = gst_pay_perc.round(4)
#     df["Margin (Billing Cost )"] = margin_value.round(2)
#     df["GST Payable (Billing Cost )"] = gst_pay_value.round(2)
#     df["Basic (Billing Cost )"] = basic_rate.round(2)
#     df["GST Reimbursement (Billing Cost ) %"] = gst_reimb_perc.round(4)
#     df["Basic Value (Billing Cost )"] = basic_val.round(2)
#     df["GST Reimbursement (Billing Cost )"] = gst_reimb_val.round(2)
#     df["Bill Value (Billing Cost )"] = bill_val.round(2)
#     return df


# # -------------------- 3) PAYMENT WORKING --------------------
# PAY_REQUIRED = [
#     "net sales (sales)",
#     "qty (sales)",
#     "gst reimbursement (billing cost )",
#     "gst reimbursement (billing cost ) %",
#     "partner name",
# ]

# def process_payment_working(df: pd.DataFrame, margin_payment_map: dict, gst_payment_threshold_map: dict) -> pd.DataFrame:
#     df = df.copy()
#     col_map = build_canon_map(df.columns)
#     if any(req not in col_map for req in PAY_REQUIRED):
#         return df

#     net_sales = to_num(df[col_map["net sales (sales)"]])
#     qty_sales = to_num(df[col_map["qty (sales)"]])
#     partner_series = df[col_map["partner name"]].fillna("")
#     partner_key = partner_series.map(normalize_partner)

#     billing_reimb_val = to_num(df[col_map["gst reimbursement (billing cost )"]])
#     billing_reimb_perc = to_num(df[col_map["gst reimbursement (billing cost ) %"]])

#     margin_perc = partner_key.map(lambda k: margin_payment_map.get(k, 0.0)).astype(float)

#     mrp_rate_pay = pd.Series(0.0, index=df.index)
#     nz = qty_sales != 0
#     mrp_rate_pay[nz] = net_sales[nz] / qty_sales[nz]

#     thresh_obj = partner_key.map(lambda k: gst_payment_threshold_map.get(k, {}))
#     pay_thresh = thresh_obj.map(lambda d: d.get("payPayThresh", 0.0)).astype(float)
#     pay_thresh = pay_thresh.where(pay_thresh != 0, 2625.0)

#     pay_low = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
#     pay_high = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
#     gst_pay_perc = pd.Series(
#         [low if rate <= thr else high for rate, thr, low, high in zip(mrp_rate_pay, pay_thresh, pay_low, pay_high)],
#         index=df.index,
#     )

#     margin_value = mrp_rate_pay * margin_perc
#     gst_pay_value = (mrp_rate_pay * gst_pay_perc / (gst_pay_perc + 1)).fillna(0)
#     basic_rate = mrp_rate_pay - margin_value - gst_pay_value

#     gst_reimb_perc = billing_reimb_perc
#     basic_val = basic_rate * qty_sales
#     gst_reimb_val = billing_reimb_val
#     bill_val = basic_val + gst_reimb_val

#     df["Margin (Payment ) %"] = margin_perc.round(4)
#     df["MRP Rate (Payment ) %"] = mrp_rate_pay.round(2)
#     df["GST Payable (Payment ) %"] = gst_pay_perc.round(4)
#     df["Margin (Payment )"] = margin_value.round(2)
#     df["GST Payable (Payment )"] = gst_pay_value.round(2)
#     df["Basic (Payment )"] = basic_rate.round(2)
#     df["GST Reimbursement (Payment ) %"] = gst_reimb_perc.round(4)
#     df["Basic Value (Payment )"] = basic_val.round(2)
#     df["GST Reimbursement (Payment )"] = gst_reimb_val.round(2)
#     df["Bill Value (Payment )"] = bill_val.round(2)
#     return df


# # -------------------- 4) RECEIVABLE RECONCILIATION --------------------
# REC_REQUIRED = [
#     "bill value (billing cost )",
#     "bill value (payment )",
#     "margin (billing cost )",
#     "gst payable (billing cost )",
#     "margin (payment )",
#     "gst payable (payment )",
#     "qty (sales)",
# ]

# def process_receivable_reconciliation(df: pd.DataFrame) -> pd.DataFrame:
#     df = df.copy()
#     col_map = build_canon_map(df.columns)
#     if any(req not in col_map for req in REC_REQUIRED):
#         return df

#     bill_b = to_num(df[col_map["bill value (billing cost )"]])
#     bill_p = to_num(df[col_map["bill value (payment )"]])
#     margin_b = to_num(df[col_map["margin (billing cost )"]])
#     gstpay_b = to_num(df[col_map["gst payable (billing cost )"]])
#     margin_p = to_num(df[col_map["margin (payment )"]])
#     gstpay_p = to_num(df[col_map["gst payable (payment )"]])
#     qty_sales = to_num(df[col_map["qty (sales)"]])

#     payable_col = (
#         col_map.get("payable")
#         or col_map.get("final payable")
#         or col_map.get("final payble")
#         or col_map.get("pay")
#         or col_map.get("final payable ")
#     )
#     dncn_col = (
#         col_map.get("dn/cn")
#         or col_map.get("final debit")
#         or col_map.get("dn")
#         or col_map.get("debit note")
#         or col_map.get("final debit ")
#     )

#     payable = to_num(df[payable_col]) if payable_col else None
#     dncn = to_num(df[dncn_col]) if dncn_col else None

#     discount_cn = bill_b - bill_p
#     df["Discount CN"] = discount_cn.round(2)

#     if payable is not None:
#         df["payment difference"] = (bill_p - payable).round(2)

#     if dncn is not None:
#         df["CN difference"] = (discount_cn - dncn).round(2)

#     df["Margin Value (Billing Rec)"] = (margin_b * qty_sales).round(2)
#     df["GST Payable (Billing Rec)"] = (gstpay_b * qty_sales).round(2)
#     df["Margin Value (Payment Rec)"] = (margin_p * qty_sales).round(2)
#     df["GST Payable (Payment Rec)"] = (gstpay_p * qty_sales).round(2)
#     return df


# # -------------------- 5) COGS --------------------
# EAN_CANDIDATES = [
#     "eancode",
#     "ean code",
#     "ean",
#     "gtin",
#     "ean/upc",
#     "ean/upc ",
#     "stock no",
#     "stockno",
#     "gs1 number",
#     "barcode scanned for billing",
#     "item code",
#     "itemcode",
#     "eancode ",
# ]

# def process_cogs(df: pd.DataFrame, cogs_map: dict) -> pd.DataFrame:
#     df = df.copy()
#     col_map = build_canon_map(df.columns)

#     ean_actual = None
#     for cand in EAN_CANDIDATES:
#         if cand in col_map:
#             ean_actual = col_map[cand]
#             break

#     if "qty" not in col_map or ean_actual is None:
#         return df

#     qty = to_num(df[col_map["qty"]])
#     ean_series = df[ean_actual].map(normalize_ean)

#     rate = ean_series.map(lambda e: cogs_map.get(e, None))
#     cogs_value = rate.copy()
#     cogs_value = cogs_value.where(cogs_value.notna(), other="not found in master")

#     rate_num = pd.to_numeric(pd.Series(rate).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
#     cogs_total = (rate_num * qty).round(2)

#     df["COGS(Value)"] = cogs_value
#     df["COGS(Rate)"] = cogs_total
#     return df


# # -------------------- 6) RAW + PIVOT --------------------
# def build_receivable_raw_and_pivot(processed_dfs):
#     raw_rows = []
#     for df in processed_dfs:
#         if df is None or df.empty:
#             continue
#         if "Partner Name" not in df.columns:
#             continue

#         tmp = pd.DataFrame({h: (df[h] if h in df.columns else "") for h in REQUIRED_HEADERS})
#         mask_any = tmp.apply(lambda r: any((v is not None) and (str(v).strip() != "") for v in r), axis=1)
#         tmp = tmp[mask_any]

#         if not tmp.empty:
#             raw_rows.append(tmp)

#     raw_df = pd.concat(raw_rows, ignore_index=True) if raw_rows else pd.DataFrame(columns=REQUIRED_HEADERS)

#     if raw_df.empty:
#         pivot_df = pd.DataFrame(columns=REQUIRED_HEADERS)
#         return raw_df, pivot_df

#     pivot_work = raw_df.copy()
#     for c in pivot_work.columns:
#         if c != "Partner Name":
#             pivot_work[c] = pd.to_numeric(pivot_work[c], errors="coerce").fillna(0)

#     pivot_df = pivot_work.groupby("Partner Name", as_index=False).sum(numeric_only=True)
#     return raw_df, pivot_df


# # -------------------- MAIN --------------------
# def main():
#     args = parse_args()

#     input_dir = Path(args.input_dir) if args.input_dir else DEFAULT_INPUT_DIR
#     output_dir = Path(args.output_dir)
#     output_dir.mkdir(parents=True, exist_ok=True)

#     output_file = output_dir / (args.output_name or DEFAULT_OUTPUT_NAME)

#     month = (args.month or "").strip().upper()
#     year = (args.year or "").strip()

#     if not input_dir.exists():
#         raise FileNotFoundError(f"Input folder not found: {input_dir}")

#     (
#         margin_billing_map,
#         margin_payment_map,
#         gst_billing_threshold_map,
#         gst_payment_threshold_map,
#         cogs_map,
#     ) = load_google_maps_and_cogs()

#     print(
#         f"Loaded maps: billing_margin={len(margin_billing_map)}, payment_margin={len(margin_payment_map)}, "
#         f"gst_billing={len(gst_billing_threshold_map)}, gst_payment={len(gst_payment_threshold_map)}, "
#         f"cogs={len(cogs_map)}"
#     )

#     # ✅ skip temporary office files (~$...) to avoid permission denied
#     excel_files = []
#     for p in list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls")):
#         if p.name.startswith("~$"):
#             continue
#         excel_files.append(p)
#     excel_files = sorted(excel_files)

#     if not excel_files:
#         print("No .xlsx or .xls files found in input folder.")
#         return

#     used_sheet_names = set()
#     tabs_written = 0
#     processed_dfs_for_pivot = []

#     with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
#         for excel_path in excel_files:
#             print(f"Reading: {excel_path.name}")
#             try:
#                 xl = pd.ExcelFile(excel_path)
#                 sheet_names = xl.sheet_names
#                 base_name = sanitize_excel_sheet_name(excel_path.stem)

#                 for sheet in sheet_names:
#                     df = pd.read_excel(excel_path, sheet_name=sheet)

#                     df = ensure_partner_name(df, excel_path.name)

#                     df = process_sales(df)
#                     df = process_billing_cost(df, margin_billing_map, gst_billing_threshold_map)
#                     df = process_payment_working(df, margin_payment_map, gst_payment_threshold_map)
#                     df = process_receivable_reconciliation(df)
#                     df = process_cogs(df, cogs_map)

#                     dest = base_name if len(sheet_names) == 1 else sanitize_excel_sheet_name(f"{base_name} - {sheet}")
#                     dest = make_unique_sheet_name(used_sheet_names, dest)

#                     df.to_excel(writer, sheet_name=dest, index=False)
#                     tabs_written += 1

#                     if dest not in EXCLUDE_SHEETS:
#                         processed_dfs_for_pivot.append(df)

#             except Exception as e:
#                 print(f"❌ Failed on {excel_path.name}: {e}")

#         raw_df, pivot_df = build_receivable_raw_and_pivot(processed_dfs_for_pivot)

#         raw_df = add_month_year_cols(raw_df, month, year)
#         pivot_df = add_month_year_cols(pivot_df, month, year)

#         raw_df.to_excel(writer, sheet_name=RAW_SHEET_NAME, index=False)
#         pivot_df.to_excel(writer, sheet_name=PIVOT_SHEET_NAME, index=False)

#     print("")
#     print("✅ Done. Master Excel created at:")
#     print(str(output_file))
#     print(f"Tabs written: {tabs_written}")
#     print(f"Receivable_Raw rows: {len(raw_df)}")
#     print(f"Final_Rec rows: {len(pivot_df)}")

#     # -------------------- PUSH ONLY Final_Rec to Google Sheet --------------------
#     load_dotenv(override=True)
#     target_sheet_id = env_clean(os.getenv("REC_TD_SHEET_ID", ""))
#     target_tab_name = env_clean(os.getenv("REC_TD_TAB_NAME", ""))

#     if not target_sheet_id or not target_tab_name:
#         print("REC_TD_SHEET_ID or REC_TD_TAB_NAME missing in .env. Skipping push.")
#         return

#     print(f"Pushing Final_Rec to Google Sheet tab '{target_tab_name}' (Month={month}, Year={year}) ...")
#     upsert_append_by_month_year(target_sheet_id, target_tab_name, pivot_df, month, year)
#     print("✅ Google Sheet update done.")


# if __name__ == "__main__":
#     main()




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

# ===================== FIX: FORCE UTF-8 PRINTING (prevents ✅/❌ crash on Windows) =====================
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ===================== DEFAULTS (used only if CLI not provided) =====================
DEFAULT_INPUT_DIR = Path(r"C:\path\to\input")
DEFAULT_OUTPUT_NAME = "MASTER_OUTPUT.xlsx"

RAW_SHEET_NAME = "Receivable_Raw"
PIVOT_SHEET_NAME = "Final_Rec"
EXCLUDE_SHEETS = {RAW_SHEET_NAME, PIVOT_SHEET_NAME, "Remark"}

REQUIRED_HEADERS = [
    "Partner Name",
    "Qty (Sales)",
    "Gross Sales (Sales)",
    "RSP (Sales)",
    "Rebate (Sales)",
    "Discount_Approved (Sales)",
    "Discount Not Approved (Sales)",
    "Total Discount (Sales)",
    "Net Sales (Sales)",
    "Basic Value (Billing Cost )",
    "GST Reimbursement (Billing Cost )",
    "Bill Value (Billing Cost )",
    "Basic Value (Payment )",
    "GST Reimbursement (Payment )",
    "Bill Value (Payment )",
    "Discount CN",
    "Margin Value (Billing Rec)",
    "GST Payable (Billing Rec)",
    "Margin Value (Payment Rec)",
    "GST Payable (Payment Rec)",
]

# ---------- Partner inference from filename ----------
PARTNER_KEYWORDS = [
    ("myntra", "Myntra jabong"),
    ("jabong", "Myntra jabong"),
    ("flipkart", "Flipkart"),
    ("v - retail", "V - Retail"),
    ("v-retail", "V - Retail"),
    ("kora", "Kora"),
    ("zuup", "ZUUP"),
    ("nirankar", "NIRANKAR"),
    ("shoppers stop", "SHOPPERS STOP"),
    ("shoppersstop", "SHOPPERS STOP"),
    ("relience centro", "Relience Centro"),
    ("reliance centro", "Relience Centro"),
    ("reliance retail ajio sor", "Reliance Retail Ajio SOR"),
    ("ajio sor", "Reliance Retail Ajio SOR"),
    ("ajio", "Reliance Retail Ajio SOR"),
    (" ls ", "LS"),
]

# -------------------- CLI --------------------
def parse_args():
    ap = argparse.ArgumentParser(description="Receivable Trade Discount processor (multi-file)")
    ap.add_argument("--input_dir", required=True, help="Folder containing input Excel files")
    ap.add_argument("--output_dir", required=True, help="Folder to write output file")
    ap.add_argument("--output_name", default=DEFAULT_OUTPUT_NAME, help="Output Excel file name")
    ap.add_argument("--month", default="", help="Month label (e.g. JAN, FEB, ...)")
    ap.add_argument("--year", default="", help="Year (e.g. 2026)")
    return ap.parse_args()


# -------------------- Common helpers --------------------
def canon(name: str) -> str:
    return " ".join(str(name).strip().lower().split())


def build_canon_map(columns):
    """
    Canonical -> actual. If duplicates exist, LAST one wins.
    """
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


def normalize_partner(name) -> str:
    return re.sub(r"[\s\-]+", "", str(name).lower()).strip()


def sanitize_excel_sheet_name(name: str) -> str:
    name = str(name).strip() or "Sheet"
    name = re.sub(r"[:\\/?*\[\]]", " ", name)
    name = re.sub(r"\s+", " ", name).strip() or "Sheet"
    return name[:31]


def make_unique_sheet_name(existing: set, desired: str) -> str:
    if desired not in existing:
        existing.add(desired)
        return desired
    i = 2
    while True:
        suffix = f"_{i}"
        base = desired[: (31 - len(suffix))]
        candidate = f"{base}{suffix}"
        if candidate not in existing:
            existing.add(candidate)
            return candidate
        i += 1


def as_decimal_percent(x):
    s = str(x).strip().replace("%", "")
    try:
        v = float(s)
    except:
        return 0.0
    return v / 100.0 if v > 1.5 else v


def get_gst_rates_from_threshold(threshold_value: float):
    if threshold_value == 2625 or threshold_value == 2500:
        return {"low": 0.05, "high": 0.18}
    return {"low": 0.12, "high": 0.18}


def normalize_ean(x) -> str:
    if x is None:
        return ""
    s = str(x).replace("\u00A0", " ").strip()
    s = re.sub(r"\.0$", "", s)
    if re.match(r"^\d+(\.\d+)?e\+\d+$", s.lower()):
        try:
            s = str(int(float(s)))
        except:
            pass
    return s


def infer_partner_from_filename(filename: str) -> str:
    f = str(filename).lower()
    f = f.replace("_", " ").replace("-", " ")
    f = " ".join(f.split())
    f_pad = f" {f} "
    for key, partner in PARTNER_KEYWORDS:
        if key.strip() == "ls":
            if " ls " in f_pad:
                return partner
        else:
            if key in f:
                return partner
    return ""


def ensure_partner_name(df: pd.DataFrame, source_filename: str) -> pd.DataFrame:
    df = df.copy()
    inferred = infer_partner_from_filename(source_filename)
    if not inferred:
        return df

    col_map = build_canon_map(df.columns)

    if "partner name" not in col_map:
        df["Partner Name"] = inferred
        return df

    partner_col = col_map["partner name"]
    s = (
        df[partner_col]
        .astype(str)
        .fillna("")
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
    )
    mask_blank = (s == "") | (s.str.lower() == "nan")
    df.loc[mask_blank, partner_col] = inferred
    return df


def env_clean(v: str, default: str = "") -> str:
    if v is None:
        return default
    s = str(v).strip()
    s = s.strip('"').strip("'").strip()
    if s.endswith(";"):
        s = s[:-1].strip()
    return s or default


# -------------------- Month/Year helper --------------------
def add_month_year_cols(df: pd.DataFrame, month: str, year: str) -> pd.DataFrame:
    df = df.copy()
    month = (month or "").strip().upper()
    year = (year or "").strip()
    if month and year:
        if "Year" in df.columns:
            df["Year"] = year
        else:
            df.insert(0, "Year", year)

        if "Month" in df.columns:
            df["Month"] = month
        else:
            df.insert(1, "Month", month)
    return df


# -------------------- Google Sheets (Service Account) --------------------
def get_gspread_client_from_env(write: bool = False):
    load_dotenv(override=True)

    b64 = os.getenv("GOOGLE_SA_JSON_B64")
    if not b64:
        raise RuntimeError("Missing GOOGLE_SA_JSON_B64 in .env (must be ONE LINE, no line breaks).")

    sa_json = json.loads(base64.b64decode(b64).decode("utf-8"))
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    if write:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]

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


def _normalize_month_year_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    if "Year" in df.columns:
        df["Year"] = df["Year"].astype(str).str.strip()
    if "Month" in df.columns:
        df["Month"] = df["Month"].astype(str).str.strip().str.upper()
    return df


def _normalize_partner_name_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    if "Partner Name" in df.columns:
        df["Partner Name"] = df["Partner Name"].astype(str).str.strip()
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


def upsert_append_by_month_year_partner(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
    """
    Final_Rec push logic:
    - match by Year + Month + Partner Name
    - remove matching existing rows
    - append new rows
    - keep unrelated rows untouched
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
        existing_df = _normalize_partner_name_for_compare(existing_df)
        df_new = _normalize_partner_name_for_compare(df_new)

        required_cols = {"Year", "Month", "Partner Name"}
        if required_cols.issubset(existing_df.columns) and required_cols.issubset(df_new.columns):
            new_keys = set(
                zip(
                    df_new["Year"].astype(str).str.strip(),
                    df_new["Month"].astype(str).str.upper(),
                    df_new["Partner Name"].astype(str).str.strip(),
                )
            )

            existing_keys = list(
                zip(
                    existing_df["Year"].astype(str).str.strip(),
                    existing_df["Month"].astype(str).str.upper(),
                    existing_df["Partner Name"].astype(str).str.strip(),
                )
            )

            mask_same = [key in new_keys for key in existing_keys]
            existing_df = existing_df.loc[[not matched for matched in mask_same]].copy()

        out_df = pd.concat([existing_df, df_new], ignore_index=True)

    if {"Year", "Month", "Partner Name"}.issubset(out_df.columns):
        out_df["Year"] = out_df["Year"].astype(str).str.strip()
        out_df["Month"] = out_df["Month"].astype(str).str.strip().str.upper()
        out_df["Partner Name"] = out_df["Partner Name"].astype(str).str.strip()
        out_df = out_df.drop_duplicates(subset=["Year", "Month", "Partner Name"], keep="last")

    out_df = out_df.where(pd.notnull(out_df), "")
    values = [out_df.columns.tolist()] + out_df.astype(object).values.tolist()

    ws.clear()
    ws.update(values, value_input_option="USER_ENTERED")


def replace_batch_by_month_year(sheet_id: str, tab_name: str, new_df: pd.DataFrame, month: str, year: str):
    """
    Receivable_Raw push logic:
    - remove all existing rows for same Year + Month
    - append new batch
    - keep other months untouched
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
        existing_df = _normalize_partner_name_for_compare(existing_df)
        df_new = _normalize_partner_name_for_compare(df_new)

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


def load_google_maps_and_cogs():
    load_dotenv(override=True)

    source_spreadsheet_id = env_clean(os.getenv("MARGIN_SPREADSHEET_ID2", ""))
    margin_sheet_name = env_clean(os.getenv("MARGIN_SHEET_NAME2", "Margine"))
    gst_threshold_sheet_name = env_clean(os.getenv("GST_THRESHOLD_SHEET_NAME2", "GST Threshold"))
    cogs_sheet_name = env_clean(os.getenv("COGS_SHEET_NAME2", "COGS Master"))

    if not source_spreadsheet_id:
        raise RuntimeError("Missing MARGIN_SPREADSHEET_ID2 in .env")

    gc = get_gspread_client_from_env(write=False)
    sh = gc.open_by_key(source_spreadsheet_id)

    # ---- Margine ----
    ws_m = sh.worksheet(margin_sheet_name)
    df_m = worksheet_to_df(ws_m)
    if df_m.empty:
        raise RuntimeError(f"No data found in sheet: {margin_sheet_name}")

    m_map = build_canon_map(df_m.columns)
    need_cols = ["partner name", "billing", "payment"]
    if any(c not in m_map for c in need_cols):
        raise RuntimeError("Margine sheet must have columns: Partner Name, Billing, Payment")

    col_partner = m_map["partner name"]
    col_billing = m_map["billing"]
    col_payment = m_map["payment"]

    margin_billing_map = {}
    margin_payment_map = {}

    for _, row in df_m.iterrows():
        p = row.get(col_partner, "")
        if not str(p).strip():
            continue
        key = normalize_partner(p)

        b = row.get(col_billing, "")
        py = row.get(col_payment, "")

        if str(b).strip():
            margin_billing_map[key] = as_decimal_percent(b)
        if str(py).strip():
            margin_payment_map[key] = as_decimal_percent(py)

    # ---- GST Threshold ----
    ws_g = sh.worksheet(gst_threshold_sheet_name)
    df_g = worksheet_to_df(ws_g)
    if df_g.empty:
        raise RuntimeError(f"No data found in sheet: {gst_threshold_sheet_name}")

    g_map = build_canon_map(df_g.columns)

    need_b = ["partner name", "gst payable (billing cost ) %", "gst reimbursement (billing cost ) %"]
    need_p = ["partner name", "gst payable (payment ) %", "gst reimbursement (payment ) %"]

    if any(c not in g_map for c in need_b):
        raise RuntimeError("GST Threshold missing Billing Cost columns")
    if any(c not in g_map for c in need_p):
        raise RuntimeError("GST Threshold missing Payment columns")

    pcol = g_map["partner name"]
    b_pay = g_map["gst payable (billing cost ) %"]
    b_rei = g_map["gst reimbursement (billing cost ) %"]
    p_pay = g_map["gst payable (payment ) %"]
    p_rei = g_map["gst reimbursement (payment ) %"]

    gst_billing_threshold_map = {}
    gst_payment_threshold_map = {}

    for _, row in df_g.iterrows():
        p = row.get(pcol, "")
        if not str(p).strip():
            continue
        key = normalize_partner(p)

        pay_b = pd.to_numeric(str(row.get(b_pay, "")).replace(",", ""), errors="coerce")
        rei_b = pd.to_numeric(str(row.get(b_rei, "")).replace(",", ""), errors="coerce")
        gst_billing_threshold_map[key] = {
            "payBillingThresh": float(pay_b) if pd.notna(pay_b) else 0.0,
            "reimbBillingThresh": float(rei_b) if pd.notna(rei_b) else 0.0,
        }

        pay_p = pd.to_numeric(str(row.get(p_pay, "")).replace(",", ""), errors="coerce")
        rei_p = pd.to_numeric(str(row.get(p_rei, "")).replace(",", ""), errors="coerce")
        gst_payment_threshold_map[key] = {
            "payPayThresh": float(pay_p) if pd.notna(pay_p) else 0.0,
            "reimbPayThresh": float(rei_p) if pd.notna(rei_p) else 0.0,
        }

    # ---- COGS Master ----
    ws_c = sh.worksheet(cogs_sheet_name)
    df_c = worksheet_to_df(ws_c)
    if df_c.empty:
        raise RuntimeError(f"No data found in sheet: {cogs_sheet_name}")

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
        rate = pd.to_numeric(str(row.get(r_col, "")).replace(",", ""), errors="coerce")
        cogs_map[ean] = float(rate) if pd.notna(rate) else 0.0

    return margin_billing_map, margin_payment_map, gst_billing_threshold_map, gst_payment_threshold_map, cogs_map


# -------------------- 1) SALES --------------------
SALES_REQUIRED = ["qty", "gross mrp", "rsp", "approved"]

def process_sales(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)
    if any(req not in col_map for req in SALES_REQUIRED):
        return df

    qty = to_num(df[col_map["qty"]])
    gross_mrp = to_num(df[col_map["gross mrp"]])
    rsp = to_num(df[col_map["rsp"]])
    approved = to_num(df[col_map["approved"]])

    qty_sales = qty
    gross_sales = gross_mrp
    rsp_sales = rsp.where(qty >= 0, rsp * -1)

    calculated_rebate = gross_sales - rsp_sales

    rebate = pd.Series(0.0, index=df.index)
    disc_approved = pd.Series(0.0, index=df.index)

    mask = approved < calculated_rebate
    rebate[mask] = 0
    disc_approved[mask] = approved[mask]

    rebate[~mask] = calculated_rebate[~mask]
    disc_approved[~mask] = approved[~mask] - rebate[~mask]

    total_discount = rebate + disc_approved
    net_sales = gross_sales - total_discount

    df["Qty (Sales)"] = qty_sales.round(0)
    df["Gross Sales (Sales)"] = gross_sales.round(2)
    df["RSP (Sales)"] = rsp_sales.round(2)
    df["Rebate (Sales)"] = rebate.round(2)
    df["Discount_Approved (Sales)"] = disc_approved.round(2)
    df["Discount Not Approved (Sales)"] = ""
    df["Total Discount (Sales)"] = total_discount.round(2)
    df["Net Sales (Sales)"] = net_sales.round(2)
    return df


# -------------------- 2) BILLING COST --------------------
BILL_REQUIRED = ["gross sales (sales)", "qty (sales)", "partner name"]

def process_billing_cost(df: pd.DataFrame, margin_billing_map: dict, gst_billing_threshold_map: dict) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)
    if any(req not in col_map for req in BILL_REQUIRED):
        return df

    gross_sales = to_num(df[col_map["gross sales (sales)"]])
    qty_sales = to_num(df[col_map["qty (sales)"]])
    partner_series = df[col_map["partner name"]].fillna("")
    partner_key = partner_series.map(normalize_partner)

    margin_perc = partner_key.map(lambda k: margin_billing_map.get(k, 0.0)).astype(float)

    mrp_rate = pd.Series(0.0, index=df.index)
    nz = qty_sales != 0
    mrp_rate[nz] = gross_sales[nz] / qty_sales[nz]

    thresh_obj = partner_key.map(lambda k: gst_billing_threshold_map.get(k, {}))
    pay_thresh = thresh_obj.map(lambda d: d.get("payBillingThresh", 0.0)).astype(float)
    reimb_thresh = thresh_obj.map(lambda d: d.get("reimbBillingThresh", 0.0)).astype(float)

    pay_thresh = pay_thresh.where(pay_thresh != 0, 1000.0)
    reimb_thresh = reimb_thresh.where(reimb_thresh != 0, pay_thresh)

    pay_low = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
    pay_high = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
    gst_pay_perc = pd.Series(
        [low if rate <= thr else high for rate, thr, low, high in zip(mrp_rate, pay_thresh, pay_low, pay_high)],
        index=df.index,
    )

    margin_value = mrp_rate * margin_perc
    gst_pay_value = (mrp_rate * gst_pay_perc / (gst_pay_perc + 1)).fillna(0)
    basic_rate = mrp_rate - margin_value - gst_pay_value

    reimb_low = reimb_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
    reimb_high = reimb_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
    gst_reimb_perc = pd.Series(
        [low if b <= thr else high for b, thr, low, high in zip(basic_rate, reimb_thresh, reimb_low, reimb_high)],
        index=df.index,
    )

    basic_val = basic_rate * qty_sales
    gst_reimb_val = basic_val * gst_reimb_perc
    bill_val = basic_val + gst_reimb_val

    df["Margin (Billing Cost ) %"] = margin_perc.round(4)
    df["MRP Rate (Billing Cost ) %"] = mrp_rate.round(2)
    df["GST Payable (Billing Cost ) %"] = gst_pay_perc.round(4)
    df["Margin (Billing Cost )"] = margin_value.round(2)
    df["GST Payable (Billing Cost )"] = gst_pay_value.round(2)
    df["Basic (Billing Cost )"] = basic_rate.round(2)
    df["GST Reimbursement (Billing Cost ) %"] = gst_reimb_perc.round(4)
    df["Basic Value (Billing Cost )"] = basic_val.round(2)
    df["GST Reimbursement (Billing Cost )"] = gst_reimb_val.round(2)
    df["Bill Value (Billing Cost )"] = bill_val.round(2)
    return df


# -------------------- 3) PAYMENT WORKING --------------------
PAY_REQUIRED = [
    "net sales (sales)",
    "qty (sales)",
    "gst reimbursement (billing cost )",
    "gst reimbursement (billing cost ) %",
    "partner name",
]

def process_payment_working(df: pd.DataFrame, margin_payment_map: dict, gst_payment_threshold_map: dict) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)
    if any(req not in col_map for req in PAY_REQUIRED):
        return df

    net_sales = to_num(df[col_map["net sales (sales)"]])
    qty_sales = to_num(df[col_map["qty (sales)"]])
    partner_series = df[col_map["partner name"]].fillna("")
    partner_key = partner_series.map(normalize_partner)

    billing_reimb_val = to_num(df[col_map["gst reimbursement (billing cost )"]])
    billing_reimb_perc = to_num(df[col_map["gst reimbursement (billing cost ) %"]])

    margin_perc = partner_key.map(lambda k: margin_payment_map.get(k, 0.0)).astype(float)

    mrp_rate_pay = pd.Series(0.0, index=df.index)
    nz = qty_sales != 0
    mrp_rate_pay[nz] = net_sales[nz] / qty_sales[nz]

    thresh_obj = partner_key.map(lambda k: gst_payment_threshold_map.get(k, {}))
    pay_thresh = thresh_obj.map(lambda d: d.get("payPayThresh", 0.0)).astype(float)
    pay_thresh = pay_thresh.where(pay_thresh != 0, 2625.0)

    pay_low = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["low"])
    pay_high = pay_thresh.map(lambda t: get_gst_rates_from_threshold(float(t))["high"])
    gst_pay_perc = pd.Series(
        [low if rate <= thr else high for rate, thr, low, high in zip(mrp_rate_pay, pay_thresh, pay_low, pay_high)],
        index=df.index,
    )

    margin_value = mrp_rate_pay * margin_perc
    gst_pay_value = (mrp_rate_pay * gst_pay_perc / (gst_pay_perc + 1)).fillna(0)
    basic_rate = mrp_rate_pay - margin_value - gst_pay_value

    gst_reimb_perc = billing_reimb_perc
    basic_val = basic_rate * qty_sales
    gst_reimb_val = billing_reimb_val
    bill_val = basic_val + gst_reimb_val

    df["Margin (Payment ) %"] = margin_perc.round(4)
    df["MRP Rate (Payment ) %"] = mrp_rate_pay.round(2)
    df["GST Payable (Payment ) %"] = gst_pay_perc.round(4)
    df["Margin (Payment )"] = margin_value.round(2)
    df["GST Payable (Payment )"] = gst_pay_value.round(2)
    df["Basic (Payment )"] = basic_rate.round(2)
    df["GST Reimbursement (Payment ) %"] = gst_reimb_perc.round(4)
    df["Basic Value (Payment )"] = basic_val.round(2)
    df["GST Reimbursement (Payment )"] = gst_reimb_val.round(2)
    df["Bill Value (Payment )"] = bill_val.round(2)
    return df


# -------------------- 4) RECEIVABLE RECONCILIATION --------------------
REC_REQUIRED = [
    "bill value (billing cost )",
    "bill value (payment )",
    "margin (billing cost )",
    "gst payable (billing cost )",
    "margin (payment )",
    "gst payable (payment )",
    "qty (sales)",
]

def process_receivable_reconciliation(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)
    if any(req not in col_map for req in REC_REQUIRED):
        return df

    bill_b = to_num(df[col_map["bill value (billing cost )"]])
    bill_p = to_num(df[col_map["bill value (payment )"]])
    margin_b = to_num(df[col_map["margin (billing cost )"]])
    gstpay_b = to_num(df[col_map["gst payable (billing cost )"]])
    margin_p = to_num(df[col_map["margin (payment )"]])
    gstpay_p = to_num(df[col_map["gst payable (payment )"]])
    qty_sales = to_num(df[col_map["qty (sales)"]])

    payable_col = (
        col_map.get("payable")
        or col_map.get("final payable")
        or col_map.get("final payble")
        or col_map.get("pay")
        or col_map.get("final payable ")
    )
    dncn_col = (
        col_map.get("dn/cn")
        or col_map.get("final debit")
        or col_map.get("dn")
        or col_map.get("debit note")
        or col_map.get("final debit ")
    )

    payable = to_num(df[payable_col]) if payable_col else None
    dncn = to_num(df[dncn_col]) if dncn_col else None

    discount_cn = bill_b - bill_p
    df["Discount CN"] = discount_cn.round(2)

    if payable is not None:
        df["payment difference"] = (bill_p - payable).round(2)

    if dncn is not None:
        df["CN difference"] = (discount_cn - dncn).round(2)

    df["Margin Value (Billing Rec)"] = (margin_b * qty_sales).round(2)
    df["GST Payable (Billing Rec)"] = (gstpay_b * qty_sales).round(2)
    df["Margin Value (Payment Rec)"] = (margin_p * qty_sales).round(2)
    df["GST Payable (Payment Rec)"] = (gstpay_p * qty_sales).round(2)
    return df


# -------------------- 5) COGS --------------------
EAN_CANDIDATES = [
    "eancode",
    "ean code",
    "ean",
    "gtin",
    "ean/upc",
    "ean/upc ",
    "stock no",
    "stockno",
    "gs1 number",
    "barcode scanned for billing",
    "item code",
    "itemcode",
    "eancode ",
]

def process_cogs(df: pd.DataFrame, cogs_map: dict) -> pd.DataFrame:
    df = df.copy()
    col_map = build_canon_map(df.columns)

    ean_actual = None
    for cand in EAN_CANDIDATES:
        if cand in col_map:
            ean_actual = col_map[cand]
            break

    if "qty" not in col_map or ean_actual is None:
        return df

    qty = to_num(df[col_map["qty"]])
    ean_series = df[ean_actual].map(normalize_ean)

    rate = ean_series.map(lambda e: cogs_map.get(e, None))
    cogs_value = rate.copy()
    cogs_value = cogs_value.where(cogs_value.notna(), other="not found in master")

    rate_num = pd.to_numeric(pd.Series(rate).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
    cogs_total = (rate_num * qty).round(2)

    df["COGS(Value)"] = cogs_value
    df["COGS(Rate)"] = cogs_total
    return df


# -------------------- 6) RAW + PIVOT --------------------
def build_receivable_raw_and_pivot(processed_dfs):
    raw_rows = []
    for df in processed_dfs:
        if df is None or df.empty:
            continue
        if "Partner Name" not in df.columns:
            continue

        tmp = pd.DataFrame({h: (df[h] if h in df.columns else "") for h in REQUIRED_HEADERS})
        mask_any = tmp.apply(lambda r: any((v is not None) and (str(v).strip() != "") for v in r), axis=1)
        tmp = tmp[mask_any]

        if not tmp.empty:
            raw_rows.append(tmp)

    raw_df = pd.concat(raw_rows, ignore_index=True) if raw_rows else pd.DataFrame(columns=REQUIRED_HEADERS)

    if raw_df.empty:
        pivot_df = pd.DataFrame(columns=REQUIRED_HEADERS)
        return raw_df, pivot_df

    pivot_work = raw_df.copy()
    for c in pivot_work.columns:
        if c != "Partner Name":
            pivot_work[c] = pd.to_numeric(pivot_work[c], errors="coerce").fillna(0)

    pivot_group_cols = []
    if "Year" in raw_df.columns:
        pivot_group_cols.append("Year")
    if "Month" in raw_df.columns:
        pivot_group_cols.append("Month")
    pivot_group_cols.append("Partner Name")

    if "Year" in raw_df.columns or "Month" in raw_df.columns:
        pivot_work = raw_df.copy()
        for c in pivot_work.columns:
            if c not in pivot_group_cols:
                pivot_work[c] = pd.to_numeric(pivot_work[c], errors="coerce").fillna(0)
        pivot_df = pivot_work.groupby(pivot_group_cols, as_index=False).sum(numeric_only=True)
    else:
        pivot_df = pivot_work.groupby("Partner Name", as_index=False).sum(numeric_only=True)

    return raw_df, pivot_df


# -------------------- MAIN --------------------
def main():
    args = parse_args()

    input_dir = Path(args.input_dir) if args.input_dir else DEFAULT_INPUT_DIR
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / (args.output_name or DEFAULT_OUTPUT_NAME)

    month = (args.month or "").strip().upper()
    year = (args.year or "").strip()

    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    (
        margin_billing_map,
        margin_payment_map,
        gst_billing_threshold_map,
        gst_payment_threshold_map,
        cogs_map,
    ) = load_google_maps_and_cogs()

    print(
        f"Loaded maps: billing_margin={len(margin_billing_map)}, payment_margin={len(margin_payment_map)}, "
        f"gst_billing={len(gst_billing_threshold_map)}, gst_payment={len(gst_payment_threshold_map)}, "
        f"cogs={len(cogs_map)}"
    )

    excel_files = []
    for p in list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls")):
        if p.name.startswith("~$"):
            continue
        excel_files.append(p)
    excel_files = sorted(excel_files)

    if not excel_files:
        print("No .xlsx or .xls files found in input folder.")
        return

    used_sheet_names = set()
    tabs_written = 0
    processed_dfs_for_pivot = []

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for excel_path in excel_files:
            print(f"Reading: {excel_path.name}")
            try:
                xl = pd.ExcelFile(excel_path)
                sheet_names = xl.sheet_names
                base_name = sanitize_excel_sheet_name(excel_path.stem)

                for sheet in sheet_names:
                    df = pd.read_excel(excel_path, sheet_name=sheet)

                    df = ensure_partner_name(df, excel_path.name)

                    df = process_sales(df)
                    df = process_billing_cost(df, margin_billing_map, gst_billing_threshold_map)
                    df = process_payment_working(df, margin_payment_map, gst_payment_threshold_map)
                    df = process_receivable_reconciliation(df)
                    df = process_cogs(df, cogs_map)

                    dest = base_name if len(sheet_names) == 1 else sanitize_excel_sheet_name(f"{base_name} - {sheet}")
                    dest = make_unique_sheet_name(used_sheet_names, dest)

                    df.to_excel(writer, sheet_name=dest, index=False)
                    tabs_written += 1

                    if dest not in EXCLUDE_SHEETS:
                        processed_dfs_for_pivot.append(df)

            except Exception as e:
                print(f"❌ Failed on {excel_path.name}: {e}")

        raw_df, pivot_df = build_receivable_raw_and_pivot(processed_dfs_for_pivot)

        raw_df = add_month_year_cols(raw_df, month, year)
        pivot_df = add_month_year_cols(pivot_df, month, year)

        raw_df.to_excel(writer, sheet_name=RAW_SHEET_NAME, index=False)
        pivot_df.to_excel(writer, sheet_name=PIVOT_SHEET_NAME, index=False)

    print("")
    print("✅ Done. Master Excel created at:")
    print(str(output_file))
    print(f"Tabs written: {tabs_written}")
    print(f"Receivable_Raw rows: {len(raw_df)}")
    print(f"Final_Rec rows: {len(pivot_df)}")

    # -------------------- PUSH Final_Rec --------------------
    load_dotenv(override=True)
    target_sheet_id = env_clean(os.getenv("REC_TD_SHEET_ID", ""))
    target_tab_name = env_clean(os.getenv("REC_TD_TAB_NAME", ""))

    if not target_sheet_id or not target_tab_name:
        print("REC_TD_SHEET_ID or REC_TD_TAB_NAME missing in .env. Skipping Final_Rec push.")
    else:
        print(
            f"Pushing Final_Rec to Google Sheet tab '{target_tab_name}' "
            f"using Year+Month+Partner Name (Month={month}, Year={year}) ..."
        )
        upsert_append_by_month_year_partner(target_sheet_id, target_tab_name, pivot_df, month, year)
        print("✅ Final_Rec Google Sheet update done.")

    # -------------------- PUSH Receivable_Raw --------------------
    base_sheet_id = env_clean(os.getenv("REC_TD_BASE_SHEET_ID", ""))
    base_tab_name = env_clean(os.getenv("REC_TD_BASE_TAB_NAME", ""))

    if not base_sheet_id or not base_tab_name:
        print("REC_TD_BASE_SHEET_ID or REC_TD_BASE_TAB_NAME missing in .env. Skipping Receivable_Raw push.")
    else:
        print(
            f"Pushing Receivable_Raw to Google Sheet tab '{base_tab_name}' "
            f"using batch replace by Year+Month (Month={month}, Year={year}) ..."
        )
        replace_batch_by_month_year(base_sheet_id, base_tab_name, raw_df, month, year)
        print("✅ Receivable_Raw Google Sheet update done.")


if __name__ == "__main__":
    main()