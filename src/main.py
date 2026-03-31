
import json
import os
import sys
import re
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
from collections import defaultdict

# Add parent directory to path so we can import config
import config


# ============================================================
# 1. AUTHENTICATE WITH GOOGLE
# ============================================================
def authenticate():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        print("[Auth] Using credentials from environment variable")
        creds_dict = json.loads(creds_json)
        credentials = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        print("[Auth] Using credentials from service_account.json")
        credentials = Credentials.from_service_account_file(
            "service_account.json", scopes=SCOPES
        )
    client = gspread.authorize(credentials)
    print("[Auth] Successfully authenticated")
    return client


# ============================================================
# 2. HELPER: Clean numeric values from sheets
# ============================================================
def clean_number(val):
    """Convert sheet cell value to float. Handles commas, dashes, blanks."""
    if val is None:
        return 0.0
    val = str(val).strip()
    if val in ("", "-", "–", "—"):
        return 0.0
    # Remove commas and whitespace
    val = val.replace(",", "").replace(" ", "")
    try:
        return float(val)
    except ValueError:
        return 0.0


# ============================================================
# 3. PULL REVENUE DATA (India + ROW)
# ============================================================
def pull_revenue(client, sheet_config):
    """
    Reads a revenue sheet and returns monthly revenue by category (Tagging).
    Returns: dict { category: { month: total_revenue } }
    
    The revenue sheets have a complex multi-row header:
    - Row 1: Section headers (BILLING FY..., REVENUE FY..., etc.)
    - Row 2: Column headers (Sr.No., Billing Unit, ..., Apr-25, May-25, etc.)
    """
    print(f"[Revenue] Reading '{sheet_config['name']}' ...")
    
    spreadsheet = client.open_by_key(sheet_config["id"])
    worksheet = spreadsheet.worksheet(sheet_config["worksheet"])
    all_data = worksheet.get_all_values()
    
    if len(all_data) < 3:
        print(f"[Revenue] WARNING: Sheet has less than 3 rows, skipping")
        return {}
    
    # Row 0 = section headers (Row 1 in sheet)
    # Row 1 = column headers (Row 2 in sheet)
    section_headers = all_data[0]
    col_headers = all_data[1]
    data_rows = all_data[2:]  # Data starts from row 3
    
    # Find the "Tagging" column index
    tagging_col = None
    for i, h in enumerate(col_headers):
        if h.strip().lower() == "tagging":
            tagging_col = i
            break
    
    if tagging_col is None:
        print("[Revenue] WARNING: Could not find 'Tagging' column")
        return {}
    
    # Find the "Revenue FY 2025-26(Estimate)" section
    # Look in section_headers for this text, then find the monthly columns under it
    estimate_start = None
    for i, h in enumerate(section_headers):
        h_clean = h.strip().lower()
        if "revenue" in h_clean and "2025-26" in h_clean and "estimate" in h_clean:
            estimate_start = i
            break
    
    if estimate_start is None:
        # Fallback: try "Revenue FY 2025-26(Budget)" 
        for i, h in enumerate(section_headers):
            h_clean = h.strip().lower()
            if "revenue" in h_clean and "2025-26" in h_clean and "budget" in h_clean:
                estimate_start = i
                break
    
    if estimate_start is None:
        print("[Revenue] WARNING: Could not find Revenue FY 2025-26 section")
        return {}
    
    print(f"[Revenue]   Revenue section starts at column index {estimate_start}")
    
    # Map month column headers to their indices within this section
    # The section has 12 monthly columns + 1 FY total column
    month_columns = {}  # { "Apr-25": col_index, ... }
    
    for i in range(estimate_start, min(estimate_start + 13, len(col_headers))):
        header = col_headers[i].strip()
        if header in config.MONTH_MAP:
            month_columns[config.MONTH_MAP[header]] = i
    
    print(f"[Revenue]   Found {len(month_columns)} month columns: {list(month_columns.keys())}")
    
    # Now aggregate revenue by Tagging and Month
    revenue = defaultdict(lambda: defaultdict(float))
    
    for row in data_rows:
        if len(row) <= tagging_col:
            continue
        tagging = row[tagging_col].strip()
        
        # Map tagging to our standard category name
        category = config.TAGGING_TO_CATEGORY.get(tagging)
        if not category:
            continue
        
        for month, col_idx in month_columns.items():
            if col_idx < len(row):
                val = clean_number(row[col_idx])
                revenue[category][month] += val
    
    # Log summary
    for cat in sorted(revenue.keys()):
        total = sum(revenue[cat].values())
        print(f"[Revenue]   {cat}: FY Total = {total:,.0f}")
    
    return revenue


# ============================================================
# 4. PULL OVERHEAD DATA
# ============================================================
def pull_overheads(client):
    """
    Reads the Overhead sheet. The sheet already has pre-computed columns:
    SaaS_India, MS_India, DSP_India, SaaS_ROW, MS_ROW, DSP_ROW
    We just sum them by month.
    """
    print(f"[Overheads] Reading '{config.OVERHEAD_SHEET['name']}' ...")
    
    spreadsheet = client.open_by_key(config.OVERHEAD_SHEET["id"])
    worksheet = spreadsheet.worksheet(config.OVERHEAD_SHEET["worksheet"])
    all_data = worksheet.get_all_values()
    
    if len(all_data) < 2:
        print("[Overheads] WARNING: Sheet has less than 2 rows")
        return {}
    
    headers = all_data[0]
    data_rows = all_data[1:]
    
    print(f"[Overheads]   {len(data_rows)} rows loaded")
    
    # Find column indices by header name
    col_map = {}
    for i, h in enumerate(headers):
        h_clean = h.strip()
        if h_clean and h_clean not in col_map:  # take first occurrence if duplicates
            col_map[h_clean] = i
    
    months_col = col_map.get("Months")
    year_col = col_map.get("Year")
    
    if months_col is None or year_col is None:
        print(f"[Overheads] WARNING: Could not find Months or Year column")
        print(f"[Overheads]   Available headers: {list(col_map.keys())[:20]}")
        return {}
    
    # Find category column indices
    cat_cols = {}
    for cat in config.CATEGORIES:
        if cat in col_map:
            cat_cols[cat] = col_map[cat]
        else:
            print(f"[Overheads]   WARNING: Column '{cat}' not found")
    
    print(f"[Overheads]   Found category columns: {list(cat_cols.keys())}")
    
    overheads = defaultdict(lambda: defaultdict(float))
    
    for row in data_rows:
        if len(row) <= max(months_col, year_col):
            continue
        
        month_name = str(row[months_col]).strip()
        year_str = str(row[year_col]).strip()
        
        try:
            year = int(float(year_str))
        except (ValueError, TypeError):
            continue
        
        month_key = config.PAYROLL_MONTH_MAP.get((month_name, year))
        if not month_key:
            continue
        
        for cat, col_idx in cat_cols.items():
            if col_idx < len(row):
                val = clean_number(row[col_idx])
                overheads[cat][month_key] += val
    
    for cat in sorted(overheads.keys()):
        total = sum(overheads[cat].values())
        print(f"[Overheads]   {cat}: FY Total = {total:,.0f}")
    
    return overheads


# ============================================================
# 5. PULL PAYROLL DATA (with special allocation rules)
# ============================================================
def pull_payroll(client, revenue_by_category):
    """
    Reads the Payroll sheet and allocates each employee's salary
    to the 6 categories based on the business rules in config.
    
    revenue_by_category is used for revenue-proportional splits.
    """
    print(f"[Payroll] Reading '{config.PAYROLL_SHEET['name']}' ...")
    
    spreadsheet = client.open_by_key(config.PAYROLL_SHEET["id"])
    worksheet = spreadsheet.worksheet(config.PAYROLL_SHEET["worksheet"])
    records = worksheet.get_all_records()
    
    df = pd.DataFrame(records)
    print(f"[Payroll]   {len(df)} rows loaded")
    
    payroll = defaultdict(lambda: defaultdict(float))
    unmatched = set()
    
    for _, row in df.iterrows():
        employee = str(row.get("Employee_Name", "")).strip()
        category_meher = str(row.get("Category_as_per_Meher", "")).strip()
        region = str(row.get("Region", "")).strip()
        month_name = str(row.get("Months", "")).strip()
        year = row.get("Year", "")
        amount = clean_number(row.get("Amount", 0))
        
        if amount == 0:
            continue
        
        try:
            year = int(float(str(year)))
        except (ValueError, TypeError):
            continue
        
        month_key = config.PAYROLL_MONTH_MAP.get((month_name, year))
        if not month_key:
            continue
        
        # --- Apply allocation rules ---
        
        # Rule: Fixed split employees (e.g., Anjum Ara)
        if employee in config.PAYROLL_FIXED_SPLIT:
            splits = config.PAYROLL_FIXED_SPLIT[employee]
            for cat, pct in splits.items():
                payroll[cat][month_key] += amount * pct
            continue
        
        # Rule: Revenue-based split employees
        if employee in config.PAYROLL_REVENUE_SPLIT:
            target_cats = config.PAYROLL_REVENUE_SPLIT[employee]
            
            # Get revenue for this month across target categories
            rev_totals = {}
            for cat in target_cats:
                rev_totals[cat] = revenue_by_category.get(cat, {}).get(month_key, 0)
            
            total_rev = sum(rev_totals.values())
            
            if total_rev > 0:
                for cat in target_cats:
                    proportion = rev_totals[cat] / total_rev
                    payroll[cat][month_key] += amount * proportion
            else:
                # If no revenue, split equally
                equal_share = amount / len(target_cats)
                for cat in target_cats:
                    payroll[cat][month_key] += equal_share
            continue
        
        # Rule: Direct mapping for everyone else
        map_key = (category_meher, region)
        category = config.PAYROLL_DIRECT_MAP.get(map_key)
        
        if category:
            payroll[category][month_key] += amount
        else:
            unmatched.add(f"{employee} ({category_meher}, {region})")
    
    if unmatched:
        print(f"[Payroll] WARNING: {len(unmatched)} unmapped employee combinations:")
        for u in sorted(unmatched):
            print(f"[Payroll]   - {u}")
    
    # Log summary
    for cat in sorted(payroll.keys()):
        total = sum(payroll[cat].values())
        print(f"[Payroll]   {cat}: FY Total = {total:,.0f}")
    
    return payroll


# ============================================================
# 6. CALCULATE P&L METRICS
# ============================================================
def calculate_metrics(revenue, overheads, payroll):
    """
    Calculates the final P&L metrics for each category and month:
    - Net Revenue (A)
    - Overhead Cost (B)
    - Payroll Cost (C)
    - Operating Income (D) = A + B + C (costs are negative)
    - Operating Income % = D / A
    - Overheads Cost % = B / A
    - Payroll Cost % = C / A
    """
    print("[Metrics] Calculating P&L metrics ...")
    
    rows = []
    
    for cat in config.CATEGORIES:
        for month in config.FY_MONTHS:
            rev = revenue.get(cat, {}).get(month, 0)
            oh = overheads.get(cat, {}).get(month, 0)
            pay = payroll.get(cat, {}).get(month, 0)
            
            # Costs should be negative (expense), revenue positive
            # Overhead and payroll amounts from sheets are positive expenses
            # Operating Income = Revenue - Overheads - Payroll
            # But the formula says D = A + B + C where B and C are costs
            # So we treat overhead and payroll as negative values
            overhead_cost = -abs(oh) if oh != 0 else 0
            payroll_cost = -abs(pay) if pay != 0 else 0
            
            operating_income = rev + overhead_cost + payroll_cost
            
            oi_pct = (operating_income / rev * 100) if rev != 0 else 0
            oh_pct = (overhead_cost / rev * 100) if rev != 0 else 0
            pay_pct = (payroll_cost / rev * 100) if rev != 0 else 0
            
            rows.append({
                "Category": cat,
                "Month": month,
                "Net_Revenue": round(rev, 2),
                "Overhead_Cost": round(overhead_cost, 2),
                "Payroll_Cost": round(payroll_cost, 2),
                "Operating_Income": round(operating_income, 2),
                "Operating_Income_Pct": f"{oi_pct:.1f}%",
                "Overhead_Cost_Pct": f"{oh_pct:.1f}%",
                "Payroll_Cost_Pct": f"{pay_pct:.1f}%",
            })
    
    # Also add FY totals per category
    for cat in config.CATEGORIES:
        total_rev = sum(revenue.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
        total_oh = sum(overheads.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
        total_pay = sum(payroll.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
        
        oh_neg = -abs(total_oh) if total_oh != 0 else 0
        pay_neg = -abs(total_pay) if total_pay != 0 else 0
        oi = total_rev + oh_neg + pay_neg
        
        oi_pct = (oi / total_rev * 100) if total_rev != 0 else 0
        oh_pct = (oh_neg / total_rev * 100) if total_rev != 0 else 0
        pay_pct = (pay_neg / total_rev * 100) if total_rev != 0 else 0
        
        rows.append({
            "Category": cat,
            "Month": "FY 2025-26 Total",
            "Net_Revenue": round(total_rev, 2),
            "Overhead_Cost": round(oh_neg, 2),
            "Payroll_Cost": round(pay_neg, 2),
            "Operating_Income": round(oi, 2),
            "Operating_Income_Pct": f"{oi_pct:.1f}%",
            "Overhead_Cost_Pct": f"{oh_pct:.1f}%",
            "Payroll_Cost_Pct": f"{pay_pct:.1f}%",
        })
    
    # Grand total across ALL categories
    for month in config.FY_MONTHS + ["FY 2025-26 Total"]:
        if month == "FY 2025-26 Total":
            total_rev = sum(
                sum(revenue.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
                for cat in config.CATEGORIES
            )
            total_oh = sum(
                sum(overheads.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
                for cat in config.CATEGORIES
            )
            total_pay = sum(
                sum(payroll.get(cat, {}).get(m, 0) for m in config.FY_MONTHS)
                for cat in config.CATEGORIES
            )
        else:
            total_rev = sum(revenue.get(cat, {}).get(month, 0) for cat in config.CATEGORIES)
            total_oh = sum(overheads.get(cat, {}).get(month, 0) for cat in config.CATEGORIES)
            total_pay = sum(payroll.get(cat, {}).get(month, 0) for cat in config.CATEGORIES)
        
        oh_neg = -abs(total_oh) if total_oh != 0 else 0
        pay_neg = -abs(total_pay) if total_pay != 0 else 0
        oi = total_rev + oh_neg + pay_neg
        
        oi_pct = (oi / total_rev * 100) if total_rev != 0 else 0
        oh_pct = (oh_neg / total_rev * 100) if total_rev != 0 else 0
        pay_pct = (pay_neg / total_rev * 100) if total_rev != 0 else 0
        
        rows.append({
            "Category": "GRAND TOTAL",
            "Month": month,
            "Net_Revenue": round(total_rev, 2),
            "Overhead_Cost": round(oh_neg, 2),
            "Payroll_Cost": round(pay_neg, 2),
            "Operating_Income": round(oi, 2),
            "Operating_Income_Pct": f"{oi_pct:.1f}%",
            "Overhead_Cost_Pct": f"{oh_pct:.1f}%",
            "Payroll_Cost_Pct": f"{pay_pct:.1f}%",
        })
    
    metrics_df = pd.DataFrame(rows)
    print(f"[Metrics] Generated {len(metrics_df)} metric rows")
    return metrics_df


# ============================================================
# 7. WRITE METRICS TO OUTPUT SHEET
# ============================================================
def write_to_output_sheet(client, metrics_df):
    print(f"[Write] Writing to '{config.OUTPUT_SHEET['name']}' ...")
    
    spreadsheet = client.open_by_key(config.OUTPUT_SHEET["id"])
    
    # Try to get or create the worksheet
    try:
        worksheet = spreadsheet.worksheet(config.OUTPUT_SHEET["worksheet"])
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=config.OUTPUT_SHEET["worksheet"],
            rows=1000,
            cols=20,
        )
        print(f"[Write]   Created new worksheet '{config.OUTPUT_SHEET['worksheet']}'")
    
    # Clear existing data
    worksheet.clear()
    print("[Write]   Cleared existing data")
    
    if not metrics_df.empty:
        # Add metadata row
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S UTC")
        
        header = metrics_df.columns.tolist()
        rows = metrics_df.values.tolist()
        rows = [[str(cell) for cell in row] for row in rows]
        
        # Write header + data
        all_rows = [
            [f"Last Updated: {timestamp}"] + [""] * (len(header) - 1),
            header,
        ] + rows
        
        worksheet.update(values=all_rows, range_name="A1")
        print(f"[Write]   Wrote {len(rows)} data rows + header")
    
    print("[Write] Output sheet updated successfully!")


# ============================================================
# 8. MAIN ENTRY POINT
# ============================================================
def main():
    print("=" * 60)
    print(f"  HectorAI Metrics Pipeline — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # Step 1: Authenticate
    client = authenticate()
    
    # Step 2: Pull revenue from India + ROW sheets
    rev_india = pull_revenue(client, config.REVENUE_INDIA_SHEET)
    rev_row = pull_revenue(client, config.REVENUE_ROW_SHEET)
    
    # Merge India + ROW revenue
    revenue = defaultdict(lambda: defaultdict(float))
    for src in [rev_india, rev_row]:
        for cat, months in src.items():
            for month, val in months.items():
                revenue[cat][month] += val
    
    print(f"\n[Revenue] Combined revenue across {len(revenue)} categories")
    for cat in sorted(revenue.keys()):
        total = sum(revenue[cat].values())
        print(f"[Revenue]   {cat}: FY Total = {total:,.0f}")
    
    # Step 3: Pull overheads
    overheads = pull_overheads(client)
    
    # Step 4: Pull payroll (needs revenue for proportional splits)
    payroll = pull_payroll(client, revenue)
    
    # Step 5: Calculate metrics
    metrics_df = calculate_metrics(revenue, overheads, payroll)
    
    # Step 6: Write to output sheet
    write_to_output_sheet(client, metrics_df)
    
    print("=" * 60)
    print("  DONE — All metrics updated successfully!")
    print("=" * 60)


if __name__ == "__main__":
    main()