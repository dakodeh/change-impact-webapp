
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.6)")

uploaded_file = st.file_uploader("Upload Change Impact Excel File", type=["xlsx"])
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = "Known Change Impacts" if "Known Change Impacts" in xls.sheet_names else xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=2)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    impact_col = next((col for col in df_raw.columns if "impact" in col.lower() and "level" in col.lower()), None)
    stakeholder_col = next((col for col in df_raw.columns if "stakeholder" in col.lower()), None)

    if not impact_col or not stakeholder_col:
        st.error("The worksheet doesn't contain recognizable Impact or Stakeholder columns.")
        st.stop()

    # Split stakeholder entries by commas and explode
    df = df_raw[[impact_col, stakeholder_col]].dropna()
    df[stakeholder_col] = df[stakeholder_col].astype(str)
    df[stakeholder_col] = df[stakeholder_col].str.split(",")
    df = df.explode(stakeholder_col)
    df[stakeholder_col] = df[stakeholder_col].str.strip()

    # Normalize impact levels
    impact_levels = ["Low", "Medium", "High"]
    df = df[df[impact_col].isin(impact_levels)]

    # Count occurrences
    df_counts = df.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0)[impact_levels]

    # Plot
    st.subheader("Change Impacts by Stakeholder Group")
    fig, ax = plt.subplots(figsize=(12, len(df_counts) * 0.5 + 2))
    df_counts.plot(kind="barh", stacked=True, ax=ax, color=["#A6CEE3", "#1F78B4", "#B2DF8A"])
    ax.set_xlabel("Number of Changes")
    ax.set_ylabel("Stakeholder Group")
    ax.set_title("Distribution of Change Impact Levels by Stakeholder")
    st.pyplot(fig)
