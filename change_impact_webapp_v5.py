
# change_impact_webapp_v15_6_1.py
# Fixes: robust column detection, prevents crash if columns not found

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

def detect_column(columns, keywords):
    for col in columns:
        col_lower = col.lower().strip()
        if all(k.lower() in col_lower for k in keywords):
            return col
    return None

def load_excel(file):
    try:
        xl = pd.ExcelFile(file)
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            if any("impact" in c.lower() for c in df.columns):
                return df
        st.warning("No sheet found with recognizable impact columns.")
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
    return None

def plot_impact(df):
    st.subheader("Degree of Impact by Stakeholder")

    impact_col = detect_column(df.columns, ["impact", "level"])
    stakeholder_col = detect_column(df.columns, ["stakeholder"])

    if not impact_col or not stakeholder_col:
        st.warning("Required columns not found for impact visualization.")
        return

    df_clean = df.dropna(subset=[impact_col, stakeholder_col])
    df_clean[impact_col] = df_clean[impact_col].str.strip().str.title()
    df_clean[stakeholder_col] = df_clean[stakeholder_col].astype(str).str.strip()

    if df_clean.empty:
        st.warning("No valid data to display for impact.")
        return

    df_counts = df_clean.groupby([stakeholder_col, df_clean[impact_col]]).size().unstack(fill_value=0)
    colors = {"Low": "green", "Medium": "orange", "High": "red"}
    df_counts.plot(kind="bar", stacked=True, color=[colors.get(c, "gray") for c in df_counts.columns])
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(plt)

def plot_perception(df):
    st.subheader("Perception of Change by Stakeholder")

    perception_col = detect_column(df.columns, ["perception"])
    stakeholder_col = detect_column(df.columns, ["stakeholder"])

    if not perception_col or not stakeholder_col:
        st.warning("Required columns not found for perception visualization.")
        return

    df_clean = df.dropna(subset=[perception_col, stakeholder_col])
    df_clean[perception_col] = df_clean[perception_col].str.strip().str.title()
    df_clean[stakeholder_col] = df_clean[stakeholder_col].astype(str).str.strip()

    if df_clean.empty:
        st.warning("No valid data to display for perception.")
        return

    df_counts = df_clean.groupby([stakeholder_col, df_clean[perception_col]]).size().unstack(fill_value=0)
    colors = {"Positive": "green", "Neutral": "blue", "Negative": "red"}
    df_counts.plot(kind="bar", stacked=True, color=[colors.get(c, "gray") for c in df_counts.columns])
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(plt)

def summarize_insights(df):
    st.subheader("Summary Insights")

    impact_col = detect_column(df.columns, ["impact", "level"])
    stakeholder_col = detect_column(df.columns, ["stakeholder"])
    perception_col = detect_column(df.columns, ["perception"])

    if not all([impact_col, stakeholder_col, perception_col]):
        st.warning("Could not generate summary insights due to missing columns.")
        return

    summary = []
    df = df.dropna(subset=[impact_col, stakeholder_col, perception_col])
    df[impact_col] = df[impact_col].str.strip().str.title()
    df[perception_col] = df[perception_col].str.strip().str.title()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.strip()

    impact_counts = df[impact_col].value_counts()
    total_changes = impact_counts.sum()
    summary.append(f"• {total_changes} changes analyzed: " +
                   ", ".join(f"{impact_counts.get(level,0)} {level}" for level in ["High", "Medium", "Low"]) + " impact")

    most_impacted = df[stakeholder_col].value_counts().head(3)
    if not most_impacted.empty:
        groups = ", ".join([f"{k} ({v} changes)" for k, v in most_impacted.items()])
        summary.append(f"• Most impacted stakeholder groups: {groups}")

    neg_perceptions = df[df[perception_col] == "Negative"]
    if not neg_perceptions.empty:
        top_neg = neg_perceptions[stakeholder_col].value_counts()
        if not top_neg.empty:
            summary.append(f"• Stakeholders with mostly negative perception: {', '.join(top_neg.index[:3])}")

    high_impact = df[df[impact_col] == "High"]
    if not high_impact.empty:
        hi_neg = high_impact[high_impact[perception_col] == "Negative"]
        group_counts = hi_neg[stakeholder_col].value_counts()
        if not group_counts.empty:
            summary.append(f"• High impact changes perceived negatively for: {', '.join(group_counts.index)}")

        multi_hi = high_impact[stakeholder_col].value_counts()
        multi_hi = multi_hi[multi_hi > 1]
        if not multi_hi.empty:
            summary.append("• Stakeholders with multiple High impact changes: " +
                           ", ".join([f"{k} ({v})" for k,v in multi_hi.items()]))

    if summary:
        st.info("📊 Insights Summary\n" + "\n".join(summary))
    else:
        st.info("No Level 2 insights could be generated.")

def main():
    st.title("Change Impact Analysis Summary Tool (v15.6.1)")
    uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])
    if uploaded_file:
        df = load_excel(uploaded_file)
        if df is not None:
            plot_impact(df)
            plot_perception(df)
            summarize_insights(df)

if __name__ == "__main__":
    main()
