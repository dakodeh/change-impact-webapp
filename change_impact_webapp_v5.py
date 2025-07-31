
# change_impact_webapp_v15_1_0.py
# Includes pie chart visualization for Potential Mitigation Strategies.

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.1.0)")

def load_data(file):
    try:
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        sheet_to_use = [name for name in sheet_names if "known change impacts" in name.lower()]
        if not sheet_to_use:
            st.error("No 'Known Change Impacts' worksheet found.")
            return None
        df = pd.read_excel(xls, sheet_name=sheet_to_use[0], header=2)
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def find_mitigation_columns(df):
    mitigation_keywords = ["Comms", "Training", "HR", "Other"]
    columns_found = {}
    for col in df.columns:
        for keyword in mitigation_keywords:
            if keyword.lower() in str(col).lower():
                columns_found[keyword] = col
    return columns_found

def plot_mitigation_pie(df, mitigation_cols):
    counts = {}
    for label, col in mitigation_cols.items():
        counts[label] = df[col].apply(lambda x: pd.notna(x) and str(x).strip() != "").sum()
    if sum(counts.values()) == 0:
        st.info("No data found for mitigation strategy pie chart.")
        return
    fig, ax = plt.subplots()
    ax.pie(counts.values(), labels=counts.keys(), autopct='%1.1f%%', startangle=90)
    ax.set_title("Distribution of Mitigation Strategies")
    st.pyplot(fig)

uploaded_file = st.file_uploader("Upload Change Impact Analysis Excel File", type=["xlsx"])
if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        mitigation_cols = find_mitigation_columns(df)
        if mitigation_cols:
            st.subheader("📊 Mitigation Strategy Distribution")
            plot_mitigation_pie(df, mitigation_cols)
        else:
            st.warning("No mitigation strategy columns found in the file.")
