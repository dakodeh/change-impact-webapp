
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.7)")

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
    perception_col = next((col for col in df_raw.columns if "perception" in col.lower()), None)
    stakeholder_col = next((col for col in df_raw.columns if "stakeholder" in col.lower()), None)

    if not impact_col or not stakeholder_col:
        st.error("The worksheet doesn't contain recognizable Impact or Stakeholder columns.")
        st.stop()

    # Clean and explode stakeholder column
    df = df_raw.dropna(subset=[impact_col, stakeholder_col]).copy()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.split(",")
    df = df.explode(stakeholder_col)
    df[stakeholder_col] = df[stakeholder_col].str.strip()

    # Normalize impact levels
    impact_levels = ["Low", "Medium", "High"]
    df = df[df[impact_col].isin(impact_levels)]

    # Count impacts
    df_counts = df.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0)[impact_levels]

    # Impact chart
    fig, ax = plt.subplots(figsize=(10, 6))
    color_map = {"Low": "#2ca02c", "Medium": "#ff7f0e", "High": "#d62728"}
    df_counts.plot(kind="barh", stacked=True, color=[color_map[imp] for imp in impact_levels], ax=ax)
    ax.set_xlabel("Number of Changes")
    ax.set_ylabel("Stakeholder Group")
    ax.set_title("Distribution of Change Impact Levels by Stakeholder")
    ax.legend(title="Level of Impact")
    st.subheader("Change Impacts by Stakeholder Group")
    st.pyplot(fig)

    # Perception chart
    if perception_col:
        df_perc = df_raw.dropna(subset=[perception_col, stakeholder_col]).copy()
        df_perc[stakeholder_col] = df_perc[stakeholder_col].astype(str).str.split(",")
        df_perc = df_perc.explode(stakeholder_col)
        df_perc[stakeholder_col] = df_perc[stakeholder_col].str.strip()

        perc_levels = ["Negative", "Neutral", "Positive"]
        df_perc = df_perc[df_perc[perception_col].isin(perc_levels)]
        df_perc_counts = df_perc.groupby([stakeholder_col, perception_col]).size().unstack(fill_value=0)[perc_levels]

        fig2, ax2 = plt.subplots(figsize=(10, 6))
        color_map2 = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
        df_perc_counts.plot(kind="barh", stacked=True, color=[color_map2[p] for p in perc_levels], ax=ax2)
        ax2.set_xlabel("Number of Changes")
        ax2.set_ylabel("Stakeholder Group")
        ax2.set_title("Perception of Change by Stakeholder")
        ax2.legend(title="Perception of Change")
        st.subheader("Perception of Change by Stakeholder")
        st.pyplot(fig2)
