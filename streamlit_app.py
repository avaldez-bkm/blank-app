import streamlit as st

import pandas as pd
from difflib import get_close_matches
import io

# Define fields to extract
KEY_METRICS_FIELDS = [
    "Scenario", "WALE", "InPlaceRent", "InPlaceNOI", "InPlaceCapRate", "DistributionsToDate",
    "CurrentLiquidity", "CostBasis", "HoldPeriod", "ExitCapRate", "ExitNOI", "ExitPrice", "ExitPricePSF",
    "LIRR", "EquityMultiple", "TotalProfit", "InitialEquity", "AdditionalEquity", "TotalEquity",
    "InitialDebt", "Holdbacks", "AdditionalProceeds", "TotalDebt", "10Y Unlevered DCF"
]

def extract_from_fmv_tab(file, filename, scenario_value):
    try:
        df = pd.read_excel(file, sheet_name="FMV", header=None)

        property_id = df.iloc[2, 0]
        version = df.iloc[3, 0]

        results = {
            "FileName": filename,
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value
        }

        labels_col_e = df.iloc[:, 4].fillna("").astype(str).str.strip().str.lower().tolist()

        for field in KEY_METRICS_FIELDS:
            if field == "Scenario":
                continue  # already set by user input
            search_term = "wale (years)" if field.lower() == "wale" else field.lower()
            match = get_close_matches(search_term, labels_col_e, n=1, cutoff=0.7)
            if match:
                idx = labels_col_e.index(match[0])
                value = df.iloc[idx, 7]  # column H
                results[field] = value
            else:
                results[field] = None

        # GAV from column K (10), FMV from column L (11), where column J (9) matches version
        version_matches = df[df.iloc[:, 9].astype(str).str.strip() == str(version)]
        if not version_matches.empty:
            idx = version_matches.index[0]
            results["GAV"] = df.iloc[idx, 10]
            results["FMV"] = df.iloc[idx, 11]
        else:
            results["GAV"] = None
            results["FMV"] = None

        return results

    except Exception as e:
        return {"FileName": filename, "Error": str(e)}

# Streamlit UI
st.title("FMV Tab Extractor for Portfolio Metrics")

scenario_input = st.text_input("Enter Scenario name to apply to all files")

uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and scenario_input:
    extracted_data = []
    for file in uploaded_files:
        extracted = extract_from_fmv_tab(file, file.name, scenario_input)
        extracted_data.append(extracted)

    df_result = pd.DataFrame(extracted_data)
    st.dataframe(df_result)

    # Download link
    csv = df_result.to_csv(index=False)
    st.download_button(
        label="Download CSV",
        data=csv,
        file_name="fmv_extracted_metrics.csv",
        mime="text/csv"
    )
