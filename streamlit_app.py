import streamlit as st
import pandas as pd
from difflib import get_close_matches

# Define fields to extract for key metrics
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

def extract_from_fmv_tab(file, filename, scenario_value):
    try:
        df = pd.read_excel(file, sheet_name="FMV", header=None)

        property_id = str(df.iloc[2, 0])
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
                continue

            search_term = "wale (years)" if field.lower() == "wale" else field.lower()
            match = get_close_matches(search_term, labels_col_e, n=1, cutoff=0.7)
            if match:
                idx = labels_col_e.index(match[0])
                value = df.iloc[idx, 7]
                results[field] = value
            else:
                results[field] = None

        dcf_match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "10y unlevered dcf"]
        if not dcf_match.empty:
            dcf_value = dcf_match.iloc[0, 1]
            results["10Y Unlevered DCF"] = dcf_value
        else:
            results["10Y Unlevered DCF"] = None

        version_matches = df[df.iloc[:, 9].astype(str).str.strip() == str(version)]
        if not version_matches.empty:
            idx = version_matches.index[0]
            results["GAV"] = df.iloc[idx, 10]
            results["FMV"] = df.iloc[idx, 11]
        else:
            results["GAV"] = None
            results["FMV"] = None

        return results, property_id, version, df

    except Exception as e:
        return {"FileName": filename, "Error": str(e)}, None, None, None

def extract_ase_cashflow(df, filename, scenario_value, property_id, version):
    try:
        dates = df.iloc[6:, 13]
        cash_flows = df.iloc[6:, 14]

        result_df = pd.DataFrame({
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value,
            "Date": dates,
            "NetCashFlowAmount": cash_flows
        }).dropna(subset=["Date", "NetCashFlowAmount"])

        return result_df

    except Exception as e:
        return pd.DataFrame([{"FileName": filename, "Error": str(e)}])

def extract_general_assumptions(df, filename, property_id, file):
    try:
        results = {"Property_ID": property_id}

        labels_col_e = df.iloc[:, 4].fillna("").astype(str).str.strip().str.lower().tolist()
        for label in GENERAL_ASSUMPTIONS_FIELDS_E:
            match = get_close_matches(label.lower(), labels_col_e, n=1, cutoff=0.7)
            if match:
                idx = labels_col_e.index(match[0])
                results[label.replace(" ", "_") + "_ID"] = df.iloc[idx, 6]
            else:
                results[label.replace(" ", "_") + "_ID"] = None

        for label in GENERAL_ASSUMPTIONS_FIELDS_A_B:
            match_rows = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == label.lower()]
            results[label.replace(" ", "_")] = match_rows.iloc[0, 1] if not match_rows.empty else None

        for label in GENERAL_ASSUMPTIONS_FIELDS_A_C:
            match_rows = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == label.lower()]
            results[label.replace(" ", "_")] = match_rows.iloc[0, 2] if not match_rows.empty else None

        fund_match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "fund"]
        results["Investment_ID"] = fund_match.iloc[0, 2] if not fund_match.empty else None

        # Additional fields from DCF Model and Inputs sheets
        dcf_df = pd.read_excel(file, sheet_name="DCF Model", header=None)
        inputs_df = pd.read_excel(file, sheet_name="Inputs", header=None)

        def get_value_from_sheet(df, col_index, label, val_col):
            matches = df[df.iloc[:, col_index].astype(str).str.strip().str.lower() == label.lower()]
            return matches.iloc[0, val_col] if not matches.empty else None

        results["Unlevered_Discount_Rate"] = get_value_from_sheet(dcf_df, 2, "Unlevered Discount Rate", 4)
        results["Terminal_Cap_Rate"] = get_value_from_sheet(dcf_df, 2, "Terminal Cap (Y-Axis) Increments", 4)
        results["Disposition_Costs"] = get_value_from_sheet(dcf_df, 2, "Disposition Costs", 4)
        results["Capital_Reserve_PSF"] = get_value_from_sheet(dcf_df, 6, "Capital Reserve", 8)
        results["Capital_Reserve_Inflation_Rate"] = get_value_from_sheet(dcf_df, 6, "Capital Reserve Inflation Rate", 8)
        results["Projected_Exit_Date"] = get_value_from_sheet(inputs_df, 7, "Exit Date", 7)
        results["Interest_Rate_Hedge"] = get_value_from_sheet(inputs_df, 7, "Interest Rate / Hedge", 8)
        results["Cap_Rate"] = get_value_from_sheet(inputs_df, 1, "Cap Rate", 2)

        return results

    except Exception as e:
        return {"FileName": filename, "Error": str(e)}

# Streamlit UI
st.title("FMV Tab Extractor for Portfolio Metrics, ASE Cash Flow, and General Assumptions")

if "scenario_input" not in st.session_state:
    st.session_state["scenario_input"] = ""

scenario_input = st.text_input("Enter Scenario name to apply to all files", st.session_state["scenario_input"])
if scenario_input != st.session_state["scenario_input"]:
    st.session_state["scenario_input"] = scenario_input

uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and st.session_state["scenario_input"]:
    extracted_data = []
    ase_cashflow_combined = pd.DataFrame()
    general_assumptions_combined = []

    for file in uploaded_files:
        st.write(f"Processing: {file.name} with Scenario: {st.session_state['scenario_input']}")
        metrics, property_id, version, df = extract_from_fmv_tab(file, file.name, st.session_state["scenario_input"])
        extracted_data.append(metrics)

        if property_id is not None and version is not None and df is not None:
            ase_df = extract_ase_cashflow(df, file.name, st.session_state["scenario_input"], property_id, version)
            ase_cashflow_combined = pd.concat([ase_cashflow_combined, ase_df], ignore_index=True)

            general_df = extract_general_assumptions(df, file.name, property_id, file)
            general_assumptions_combined.append(general_df)

    df_result = pd.DataFrame(extracted_data)
    st.subheader("Key Metrics Extracted")
    st.dataframe(df_result)

    csv_main = df_result.to_csv(index=False)
    st.download_button(
        label="Download Key Metrics CSV",
        data=csv_main,
        file_name="fmv_extracted_metrics.csv",
        mime="text/csv"
    )

    if not ase_cashflow_combined.empty:
        st.subheader("ASE Cash Flow Extracted")
        st.dataframe(ase_cashflow_combined)

        csv_cashflow = ase_cashflow_combined.to_csv(index=False)
        st.download_button(
            label="Download ASE Cash Flow CSV",
            data=csv_cashflow,
            file_name="ase_cashflow_extracted.csv",
            mime="text/csv"
        )

    if general_assumptions_combined:
        df_general = pd.DataFrame(general_assumptions_combined)
        st.subheader("General Assumptions Extracted")
        st.dataframe(df_general)

        csv_general = df_general.to_csv(index=False)
        st.download_button(
            label="Download General Assumptions CSV",
            data=csv_general,
            file_name="general_assumptions.csv",
            mime="text/csv"
        )
