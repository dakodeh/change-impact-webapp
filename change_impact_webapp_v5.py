# change_impact_webapp_v5.py
# Tool Version: v15.1.5
# Last updated: 2025-08-01 15:45

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.1.5)")
st.markdown("**Tool Version**: v15.1.5  
_Last updated: 2025-08-01 15:45_", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload a completed Change Impact Assessment", type=[".xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Known Change Impacts", header=2)
    except Exception as e:
        st.error(f"Failed to read worksheet: {e}")
        st.stop()

    mitigation_columns = []
    for col in df.columns:
        if isinstance(col, str) and col.strip().lower() in ["comms", "training", "hr", "other"]:
            mitigation_columns.append(col)

    if mitigation_columns:
        mitigation_counts = {}
        for col in mitigation_columns:
            mitigation_counts[col] = df[col].apply(lambda x: 1 if pd.notna(x) and str(x).strip() != "" else 0).sum()

        if sum(mitigation_counts.values()) > 0:
            fig, ax = plt.subplots(figsize=(6, 6))
            wedges, texts, autotexts = ax.pie(
                mitigation_counts.values(),
                labels=mitigation_counts.keys(),
                autopct="%1.1f%%",
                startangle=140,
                explode=[0.1 if v == max(mitigation_counts.values()) else 0 for v in mitigation_counts.values()],
                wedgeprops=dict(width=0.4),
                textprops=dict(color="black")
            )
            ax.set_title("Distribution of Mitigation Strategies")
            st.pyplot(fig)
        else:
            st.warning("No data found for mitigation strategy pie chart.")
    else:
        st.warning("No recognizable mitigation strategy columns found.")

    st.markdown("---")
    st.info("✅ Placeholder: Additional charts and summary insights logic would appear here (e.g., Change Impacts by Stakeholder, Perception, etc.)")
else:
    st.info("📤 Please upload a completed assessment file to begin analysis.")
