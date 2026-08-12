#!/usr/bin/env python3
"""
FFIEC Peer Report Generator - FNBO Test
Self-contained: edit the 4 variables below, then run.
"""

import os
import sys
from datetime import datetime

import pandas as pd
from ffiec_data_connect import OAuth2Credentials, collect_data
from ffiec_data_connect.exceptions import CredentialError, NoDataError, RateLimitError


# ═════════════════════════════════════════════════════════════════════════════
# EDIT THESE 4 VARIABLES FOR YOUR SETUP
# ═════════════════════════════════════════════════════════════════════════════
BANK_NAME = "First National Bank of Omaha"
RSSD_ID = "527954"
REPORTING_PERIOD = "03/31/2026"   # Format: MM/DD/YYYY
OUTPUT_FILE = "fnbo_peer_report.xlsx"
# ═════════════════════════════════════════════════════════════════════════════


# 30 most important peer fields (MDRM code → human name)
PEER_FIELDS = {
    "RCON2170": "Total Assets",
    "RCON2200": "Total Deposits",
    "RCON3210": "Total Equity Capital",
    "RCON2122": "Total Loans & Leases, Net",
    "RCON1754": "Total Investment Securities",
    "RCON0010": "Cash & Due from Banks",
    "RCON3190": "Total Borrowings",
    "RCON3123": "Allowance for Loan & Lease Losses",
    "RCON3545": "Trading Assets",
    "RCON2160": "Other Assets",
    "RIAD4010": "Total Interest Income",
    "RIAD4073": "Total Interest Expense",
    "RIAD4074": "Net Interest Income",
    "RIAD4079": "Total Noninterest Income",
    "RIAD4093": "Total Noninterest Expense",
    "RIAD4230": "Provision for Loan & Lease Losses",
    "RIAD4301": "Income Before Taxes",
    "RIAD4340": "Net Income",
    "RCON1403": "Nonaccrual Loans",
    "RCON1407": "Loans 30-89 Days Past Due",
    "RCON1408": "Loans 90+ Days Past Due",
    "RIAD4635": "Net Charge-offs",
    "RCONF180": "Troubled Debt Restructurings",
    "RCON8274": "Tier 1 Capital",
    "RCON3792": "Total Risk-Based Capital",
    "RCON2232": "Risk-Weighted Assets",
    "RCONP859": "Common Equity Tier 1",
    "RCON1410": "Commercial & Industrial Loans",
    "RCON1420": "Total Real Estate Loans",
    "RCON1797": "Total Consumer Loans",
}


def fetch_bank(creds, reporting_period, rssd_id):
    try:
        df = collect_data(
            creds,
            reporting_period=reporting_period,
            rssd_id=rssd_id,
            series="call",
            output_type="pandas",
            force_null_types="pandas",
        )
        return df
    except NoDataError:
        print(f"  No data found for RSSD {rssd_id}")
        return pd.DataFrame()
    except RateLimitError as e:
        print(f"  Rate limited: retry after {e.retry_after}s")
        raise
    except CredentialError as e:
        print(f"  Authentication failed: {e}")
        raise


def filter_peer_fields(df):
    if df.empty:
        return df
    matched = {}
    unmatched = []
    for mdrm, label in PEER_FIELDS.items():
        if mdrm in df.columns:
            matched[mdrm] = df[mdrm].iloc[0] if not df[mdrm].empty else None
        else:
            hits = [c for c in df.columns if mdrm.lower() in c.lower()]
            if hits:
                matched[mdrm] = df[hits[0]].iloc[0] if not df[hits[0]].empty else None
            else:
                unmatched.append(mdrm)
    result = pd.DataFrame([{PEER_FIELDS[k]: v for k, v in matched.items()}])
    if unmatched:
        print(f"  Could not match {len(unmatched)} fields")
    return result


def compute_ratios(raw_df):
    def get(col):
        if col in raw_df.columns:
            val = raw_df[col]
            if hasattr(val, "iloc"):
                val = val.iloc[0]
            try:
                return float(val)
            except (TypeError, ValueError):
                return 0.0
        return 0.0

    assets = get("RCON2170")
    equity = get("RCON3210")
    loans = get("RCON2122")
    deposits = get("RCON2200")
    nii = get("RIAD4074")
    nonii = get("RIAD4079")
    nonie = get("RIAD4093")
    net_income = get("RIAD4340")
    nonaccrual = get("RCON1403")
    pastdue_30_89 = get("RCON1407")
    pastdue_90 = get("RCON1408")
    llr = get("RCON3123")
    chargeoffs = get("RIAD4635")
    tdr = get("RCONF180")
    cet1 = get("RCONP859")
    tier1 = get("RCON8274")
    total_rbc = get("RCON3792")
    rwa = get("RCON2232")
    cre = get("RCON1420")
    cni = get("RCON1410")

    def safe_div(n, d):
        return round(n / d, 4) if d and d != 0 else None

    ratios = {
        "ROA": safe_div(net_income, assets),
        "ROE": safe_div(net_income, equity),
        "NIM": safe_div(nii, assets),
        "Efficiency Ratio": safe_div(nonie, (nii + nonii)) if (nii + nonii) != 0 else None,
        "Equity / Assets": safe_div(equity, assets),
        "Loans / Deposits": safe_div(loans, deposits),
        "LLR / Loans": safe_div(llr, loans),
        "NPL Ratio": safe_div(nonaccrual, loans),
        "Past Due 30-89 / Loans": safe_div(pastdue_30_89, loans),
        "Past Due 90+ / Loans": safe_div(pastdue_90, loans),
        "Net Charge-off Ratio": safe_div(chargeoffs, loans),
        "TDR / Loans": safe_div(tdr, loans),
        "CET1 Ratio": safe_div(cet1, rwa),
        "Tier 1 RBC Ratio": safe_div(tier1, rwa),
        "Total RBC Ratio": safe_div(total_rbc, rwa),
        "CRE / Capital": safe_div(cre, equity),
        "C&I / Total Loans": safe_div(cni, loans),
    }
    return pd.DataFrame([ratios])


def write_excel(peer_df, ratio_df, filename, bank_name, period):
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        peer_df.to_excel(writer, sheet_name="Peer Fields", index=False)
        _autofit(writer.sheets["Peer Fields"])

        ratio_df.to_excel(writer, sheet_name="Key Ratios", index=False)
        _autofit(writer.sheets["Key Ratios"])

        meta = pd.DataFrame({
            "Bank": [bank_name],
            "Reporting Period": [period],
            "Generated": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        })
        meta.to_excel(writer, sheet_name="Metadata", index=False)
        _autofit(writer.sheets["Metadata"])
    print(f"Saved: {filename}")


def _autofit(worksheet):
    for column in worksheet.columns:
        max_length = 0
        col_letter = column[0].column_letter
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)


def main():
    username = os.environ.get("FFIEC_USERNAME")
    token = os.environ.get("FFIEC_BEARER_TOKEN")

    if not username or not token:
        print("Error: FFIEC_USERNAME and FFIEC_BEARER_TOKEN env vars required.")
        sys.exit(1)

    creds = OAuth2Credentials(username=username, bearer_token=token)

    if creds.is_expired:
        print("Error: FFIEC token expired. Generate a new one at cdr.ffiec.gov")
        sys.exit(1)

    print(f"Pulling {BANK_NAME} (RSSD: {RSSD_ID}) for {REPORTING_PERIOD}...")
    print(f"Token expires: {creds.token_expires}")

    raw_df = fetch_bank(creds, REPORTING_PERIOD, RSSD_ID)

    if raw_df.empty:
        print("Error: No data retrieved.")
        sys.exit(1)

    print(f"Raw data: {len(raw_df.columns)} columns")
    print("Sample columns:", list(raw_df.columns)[:10])

    peer_df = filter_peer_fields(raw_df)
    ratio_df = compute_ratios(raw_df)

    write_excel(peer_df, ratio_df, OUTPUT_FILE, BANK_NAME, REPORTING_PERIOD)
    print(f"Done. Peer fields: {len(peer_df.columns)} | Ratios: {len(ratio_df.columns)}")


if __name__ == "__main__":
    main()
