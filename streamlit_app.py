import streamlit as st
import pandas as pd
from difflib import get_close_matches

# Field Constants
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

# Column Index Constants
LABEL_COL = 4  # Column E
VALUE_COL = 7  # Column H


def extract_from_fmv_tab(df, filename, scenario_value):
    try:
        property_id = str(df.iloc[2, 0])
        version = df.iloc[3, 0]

        results = {
            "FileName": filename,
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value
        }

        labels_col = df.iloc[:, LABEL_COL].fillna("").astype(str).str.strip().str.lower().tolist()

        for field in [f for f in KEY_METRICS_FIELDS if f != "Scenario"]:
            search_term = "wale (years)" if field.lower() == "wale" else field.lower()
            match = get_close_matches(search_term, labels_col, n=1, cutoff=0.7)
            if match:
                idx = labels_col.index(match[0])
                results[field] = df.iloc[idx, VALUE_COL]
            else:
                st.warning(f"Could not find match for key metric field: {field}")
                results[field] = None

        dcf_match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "10y unlevered dcf"]
        results["10Y Unlevered DCF"] = dcf_match.iloc[0, 1] if not dcf_match.empty else None

        version_matches = df[df.iloc[:, 9].astype(str).str.strip() == str(version)]
        if not version_matches.empty:
            idx = version_matches.index[0]
            results["GAV"] = df.iloc[idx, 10]
            results["FMV"] = df.iloc[idx, 11]
        else:
            results["GAV"] = results["FMV"] = None

        return results, property_id, version, df

    except Exception as e:
        return {"FileName": filename, "Error": f"[extract_from_fmv_tab] {str(e)}"}, None, None, None


def extract_ase_cashflow(df, filename, scenario_value, property_id, version):
    try:
        result_df = pd.DataFrame({
            "Property_ID": property_id,
            "Version": version,
            "Scenario": scenario_value,
            "Date": df.iloc[6:, 13],
            "NetCashFlowAmount": df.iloc[6:, 14]
        }).dropna(subset=["Date", "NetCashFlowAmount"])

        return result_df
    except Exception as e:
        return pd.DataFrame([{"FileName": filename, "Error": f"[extract_ase_cashflow] {str(e)}"}])


def extract_general_assumptions(df, dcf_df, inputs_df, filename, property_id):
    try:
        results = {"Property_ID": property_id}

        labels_col_e = df.iloc[:, LABEL_COL].fillna("").astype(str).str.strip().str.lower().tolist()
        for label in GENERAL_ASSUMPTIONS_FIELDS_E:
            match = get_close_matches(label.lower(), labels_col_e, n=1, cutoff=0.7)
            if match:
                idx = labels_col_e.index(match[0])
                results[label.replace(" ", "_") + "_ID"] = df.iloc[idx, 6]
            else:
                st.warning(f"Could not find match for: {label}")
                results[label.replace(" ", "_") + "_ID"] = None

        labels_col_a = df.iloc[:, 0].fillna("").astype(str).str.strip().str.lower().tolist()
        for label in GENERAL_ASSUMPTIONS_FIELDS_A_B:
            match = get_close_matches(label.lower(), labels_col_a, n=1, cutoff=0.7)
            if match:
                idx = labels_col_a.index(match[0])
                results[label.replace(" ", "_")] = df.iloc[idx, 1]
            else:
                st.warning(f"Could not find match for: {label}")
                results[label.replace(" ", "_")] = None

        for label in GENERAL_ASSUMPTIONS_FIELDS_A_C:
            match = get_close_matches(label.lower(), labels_col_a, n=1, cutoff=0.7)
            if match:
                idx = labels_col_a.index(match[0])
                results[label.replace(" ", "_")] = df.iloc[idx, 2]
            else:
                st.warning(f"Could not find match for: {label}")
                results[label.replace(" ", "_")] = None

        fund_match = df[df.iloc[:, 0].astype(str).str.strip().str.lower() == "fund"]
        results["Investment_ID"] = fund_match.iloc[0, 2] if not fund_match.empty else None

        def get_value_from_sheet(df, col_index, label, val_col):
            labels = df.iloc[:, col_index].fillna("").astype(str).str.strip().str.lower().tolist()
            match = get_close_matches(label.lower(), labels, n=1, cutoff=0.7)
            if match:
                idx = labels.index(match[0])
                return df.iloc[idx, val_col]
            return None

        results["Unlevered_Discount_Rate"] = get_value_from_sheet(dcf_df, 2, "Unlevered Discount Rate", 4)
        results["Terminal_Cap_Rate"] = get_value_from_sheet(dcf_df, 2, "Terminal Cap (Y-Axis) Increments", 4)
        results["Disposition_Costs"] = get_value_from_sheet(dcf_df, 2, "Disposition Costs", 4)
        results["Capital_Reserve_PSF"] = get_value_from_sheet(dcf_df, 6, "Capital Reserve", 8)
        results["Capital_Reserve_Inflation_Rate"] = get_value_from_sheet(dcf_df, 6, "Capital Reserve Inflation Rate", 8)
        results["Interest_Rate_Hedge"] = get_value_from_sheet(inputs_df, 7, "Interest Rate / Hedge", 8)
        results["Cap_Rate"] = get_value_from_sheet(inputs_df, 1, "Cap Rate", 2)

        labels_col_h = inputs_df.iloc[:, 7].fillna("").astype(str).str.strip().str.lower().tolist()
        match = get_close_matches("exit date", labels_col_h, n=1, cutoff=0.7)
        if match:
            idx = labels_col_h.index(match[0])
            results["Projected_Exit_Date"] = inputs_df.iloc[idx, 8]
        else:
            results["Projected_Exit_Date"] = None

        return results

    except Exception as e:
        return {"FileName": filename, "Error": f"[extract_general_assumptions] {str(e)}"}


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

    progress_bar = st.progress(0)
    percent_text = st.empty()
    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):
        xls = pd.ExcelFile(file)

        try:
            df = xls.parse("FMV", header=None)
            dcf_df = xls.parse("DCF Model", header=None)
            inputs_df = xls.parse("Inputs", header=None)

            metrics, property_id, version, df = extract_from_fmv_tab(df, file.name, st.session_state["scenario_input"])
            extracted_data.append(metrics)

            if property_id and version:
                ase_df = extract_ase_cashflow(df, file.name, st.session_state["scenario_input"], property_id, version)
                ase_cashflow_combined = pd.concat([ase_cashflow_combined, ase_df], ignore_index=True)

                general_df = extract_general_assumptions(df, dcf_df, inputs_df, file.name, property_id)
                general_assumptions_combined.append(general_df)

        except Exception as e:
            st.error(f"Failed to process {file.name}: {str(e)}")

        progress = int((i + 1) / total_files * 100)
        progress_bar.progress((i + 1) / total_files)
        percent_text.markdown(f"**Progress: {progress}%**")

    df_result = pd.DataFrame(extracted_data)
    st.subheader("Key Metrics Extracted")
    st.dataframe(df_result)
    st.download_button("Download Key Metrics CSV", df_result.to_csv(index=False), "fmv_extracted_metrics.csv", "text/csv")

    if not ase_cashflow_combined.empty:
        st.subheader("ASE Cash Flow Extracted")
        st.dataframe(ase_cashflow_combined)
        st.download_button("Download ASE Cash Flow CSV", ase_cashflow_combined.to_csv(index=False), "ase_cashflow_extracted.csv", "text/csv")

    if general_assumptions_combined:
        df_general = pd.DataFrame(general_assumptions_combined)
        st.subheader("General Assumptions Extracted")
        st.dataframe(df_general)
        st.download_button("Download General Assumptions CSV", df_general.to_csv(index=False), "general_assumptions.csv", "text/csv")
