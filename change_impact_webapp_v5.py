
# change_impact_webapp_v15_0_4.py
# Updated to robustly detect 'Level of Impact' and 'Perception of Change' columns regardless of position or name variation

import streamlit as st
import pandas as pd

def find_column(columns, keywords):
    for col in columns:
        if all(k.lower() in col.lower() for k in keywords):
            return col
    return None

def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]  # assume first sheet
    df = xls.parse(sheet_name)
    df.columns = df.columns.str.strip()  # Strip whitespace
    return df

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.4)")

uploaded_file = st.file_uploader("Upload your Change Impact Excel file", type=["xlsx"])
if uploaded_file:
    df = load_data(uploaded_file)

    impact_col = find_column(df.columns, ["Impact", "Level"])
    perception_col = find_column(df.columns, ["Perception", "Change"])
    stakeholder_col = find_column(df.columns, ["Stakeholder"])

    if not impact_col or not perception_col:
        st.error("The worksheet doesn't contain recognizable Impact or Perception columns.")
    else:
        st.success(f"Columns detected - Impact: '{impact_col}', Perception: '{perception_col}'")
        st.write(df[[impact_col, perception_col, stakeholder_col]].head())
