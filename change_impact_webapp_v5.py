
# Streamlit Change Impact Web App - v15.0.9
# Adds summary insights back in with bullet points.

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
__version__ = '15.1.3'
__last_updated__ = '2025-08-01 11:35'
st.title("Change Impact Analysis Summary Tool (v15.0.9)")
st.markdown(f"**Tool Version**: v{__version__}")
st.markdown(f"_Last updated: {__last_updated__}_", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Change Impact Excel File", type=["xlsx"])
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = "Known Change Impacts" if "Known Change Impacts" in xls.sheet_names else xls.sheet_names[0]
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=2)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    impact_col = next((c for c in df_raw.columns if "impact" in c.lower() and "level" in c.lower()), None)
    perception_col = next((c for c in df_raw.columns if "perception" in c.lower()), None)
    stakeholder_col = next((c for c in df_raw.columns if "stakeholder" in c.lower()), None)

    if not impact_col or not stakeholder_col:
        st.error("Could not detect required columns in worksheet.")
        st.stop()

    # Expand comma‑separated stakeholders
    df = df_raw.dropna(subset=[impact_col, stakeholder_col]).copy()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.split(',')
    df = df.explode(stakeholder_col)
    df[stakeholder_col] = df[stakeholder_col].str.strip()

    ### --- Visualization helpers ---
    def horiz_stacked(data, value_col, title, colors, legend_title):
        counts = data.groupby([stakeholder_col, value_col]).size().unstack(fill_value=0)
        counts = counts[[c for c in colors.keys() if c in counts.columns]]
        fig, ax = plt.subplots(figsize=(11, max(4, len(counts) * 0.45)))
        left = [0]*len(counts)
        for cat in counts.columns:
            vals = counts[cat]
            ax.barh(counts.index, vals, left=left, color=colors[cat], label=cat)
            # add labels
            for i, (v,lft) in enumerate(zip(vals, left)):
                if v > 0:
                    ax.text(lft+v/2, i, int(v), ha='center', va='center', color='white', fontsize=8)
            left = left + vals
        ax.set_xlabel("Number of Changes")
        ax.set_ylabel("Stakeholder Group")
        ax.set_title(title)
        ax.legend(title=legend_title, bbox_to_anchor=(1.05,1), loc='upper left')
        st.pyplot(fig)
        plt.clf()

    st.subheader("Change Impacts by Stakeholder Group")
    impact_colors = {"Low":"green","Medium":"orange","High":"red"}
    horiz_stacked(df, impact_col, "Distribution of Change Impact Levels by Stakeholder", impact_colors, "Level of Impact")

    if perception_col:
        st.subheader("Perception of Change by Stakeholder")
        perception_colors = {"Positive":"green","Neutral":"blue","Negative":"red"}
        horiz_stacked(df, perception_col, "Distribution of Change Perception Levels by Stakeholder", perception_colors, "Perception of Change")

    ### --- Summary Insights ---
    st.subheader("Summary Insights")

    bullets = []

    # total counts
    total_changes = len(df)
    impact_counts = df[impact_col].value_counts().to_dict()
    bullets.append(f"• **{total_changes}** total changes analyzed: " + ", ".join(f"{lvl} **{cnt}**" for lvl,cnt in impact_counts.items()))

    # most impacted stakeholder groups (top 3)
    top_stake = df[stakeholder_col].value_counts().head(3)
    bullets.append("• Most impacted stakeholder groups: " + ", ".join(f"{grp} ({cnt})" for grp,cnt in top_stake.items()))

    # mostly negative perception
    if perception_col:
        perc = df.groupby(stakeholder_col)[perception_col].apply(lambda s: (s=='Negative').mean())
        neg_groups = perc[perc>=0.6].index.tolist()
        if neg_groups:
            bullets.append("• Stakeholders with mostly negative perception: " + ", ".join(neg_groups))

    # high impact negative
    high_neg = []
    if perception_col:
        mask = (df[impact_col]=='High') & (df[perception_col]=='Negative')
        high_neg = df.loc[mask, stakeholder_col].unique().tolist()
        if high_neg:
            bullets.append("• High‑impact changes perceived negatively for: " + ", ".join(high_neg))

    # multiple high impacts
    high_counts = df[df[impact_col]=='High'][stakeholder_col].value_counts()
    multi_high = high_counts[high_counts>1]
    if not multi_high.empty:
        bullets.append("• Stakeholders with multiple High impact changes: " + ", ".join(multi_high.index))

    if bullets:
        for b in bullets:
            st.markdown(b)
    else:
        st.write("No noteworthy patterns detected.")
