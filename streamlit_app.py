import streamlit as st
import pandas as pd
from difflib import get_close_matches

# Constants
KEY_METRICS_FIELDS = [
    "Scenario", "WALE", "InPlaceRent", "InPlaceNOI", "InPlaceCapRate", "DistributionsToDate",
    "CurrentLiquidity", "CostBasis", "HoldPeriod", "ExitCapRate", "ExitNOI", "ExitPrice", "ExitPricePSF",
    "LIRR", "EquityMultiple", "TotalProfit", "InitialEquity", "AdditionalEquity", "TotalEquity",
    "InitialDebt", "Holdbacks", "AdditionalProceeds", "TotalDebt"
]

GENERAL_ASSUMPTIONS_FIELDS_E = ["Product Type"]
GENERAL_ASSUMPTIONS_FIELDS_A_B = ["Interest Rate Caps"]
GENERAL_ASSUMPTIONS_FIELDS_A_C = [
    "Valuation Basis", "Valuation Methodology", "Levered Discount Rate",
    "Stage of Asset Life", "Appraised Value"
]

def fuzzy_find(df, column_idx, label, value_col, filename, missing_list, label_alias=None):
    labels = df.iloc[:, column_idx].fillna("").astype(str).str.strip().str.lower().tolist()
    search_term = (label_alias or label).lower()
    match = get_close_matches(search_term, labels, n=1, cutoff=0.7)
    if match:
        return df.iloc[labels.index(match[0]), value_col]
    else:
        missing_list.append(f"[{filename}] Missing: {label}")
        return None

def extract_from_fmv_tab(file, filename, scenario_value):
    try:
        df = pd.read_excel(file, sheet_name="FMV", header=None)

        # Property ID and version (magic numbers clarified)
        property_id = str(df.iloc[2, 0])  # Row 3, Col A
        version = df.iloc[3, 0]           # Row 4, Col A

        results = {
            "FileName": filename,
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value
        }

        missing_labels = []
        labels_col_e = df.iloc[:, 4].fillna("").astype(str).str.strip().str.lower().tolist()

        for field in KEY_METRICS_FIELDS:
            if field == "Scenario":
                continue

            search_term = "wale (years)" if field.lower() == "wale" else field.lower()
            match = get_close_matches(search_term, labels_col_e, n=1, cutoff=0.7)
            if match:
                idx = labels_col_e.index(match[0])
                results[field] = df.iloc[idx, 7]
            else:
                results[field] = None
                missing_labels.append(f"[{filename}] Missing: {field} (search: '{search_term}')")

        dcf_match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "10y unlevered dcf"]
        results["10Y Unlevered DCF"] = dcf_match.iloc[0, 1] if not dcf_match.empty else None
        if dcf_match.empty:
            missing_labels.append(f"[{filename}] Missing: 10Y Unlevered DCF")

        version_matches = df[df.iloc[:, 9].astype(str).str.strip() == str(version)]
        if not version_matches.empty:
            idx = version_matches.index[0]
            results["GAV"] = df.iloc[idx, 10]
            results["FMV"] = df.iloc[idx, 11]
        else:
            results["GAV"] = None
            results["FMV"] = None
            missing_labels.append(f"[{filename}] Missing: GAV / FMV for version {version}")

        return results, property_id, version, df, missing_labels

    except Exception as e:
        return {"FileName": filename, "Error": str(e)}, None, None, None, [f"{filename} extraction failed: {str(e)}"]

def extract_ase_cashflow(df, filename, scenario_value, property_id, version):
    try:
        dates = df.iloc[6:, 13]  # Row 7+, Col N
        cash_flows = df.iloc[6:, 14]  # Row 7+, Col O

        result_df = pd.DataFrame({
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value,
            "Date": pd.to_datetime(dates, errors='coerce'),
            "NetCashFlowAmount": cash_flows
        }).dropna(subset=["Date", "NetCashFlowAmount"])

        return result_df

    except Exception as e:
        return pd.DataFrame([{"FileName": filename, "Error": str(e)}])
