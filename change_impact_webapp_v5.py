
# change_impact_webapp_v15_6.py
# Fully functional version with Level 2 Summary Insights logic embedded

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

def load_excel(file):
    try:
        xl = pd.ExcelFile(file)
        if "Known Change Impacts" in xl.sheet_names:
            df = xl.parse("Known Change Impacts")
        else:
            df = xl.parse(xl.sheet_names[0])
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def plot_impact(df):
    st.subheader("Degree of Impact by Stakeholder")
    impact_col = [col for col in df.columns if "Impact" in col and "Level" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]
    df_counts = df.groupby([stakeholder_col, df[impact_col].str.title()]).size().unstack(fill_value=0)
    colors = {"Low": "green", "Medium": "orange", "High": "red"}
    df_counts.plot(kind="bar", stacked=True, color=[colors.get(c, "gray") for c in df_counts.columns])
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(plt)

def plot_perception(df):
    st.subheader("Perception of Change by Stakeholder")
    perception_col = [col for col in df.columns if "Perception" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]
    df_counts = df.groupby([stakeholder_col, df[perception_col].str.title()]).size().unstack(fill_value=0)
    colors = {"Positive": "green", "Neutral": "blue", "Negative": "red"}
    df_counts.plot(kind="bar", stacked=True, color=[colors.get(c, "gray") for c in df_counts.columns])
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(plt)

def summarize_insights(df):
    st.subheader("Summary Insights")
    impact_col = [col for col in df.columns if "Impact" in col and "Level" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]
    perception_col = [col for col in df.columns if "Perception" in col][0]

    summary = []

    impact_counts = df[impact_col].value_counts()
    total_changes = impact_counts.sum()
    summary.append(f"• {total_changes} changes analyzed: " +
                   ", ".join(f"{impact_counts.get(level,0)} {level}" for level in ["High", "Medium", "Low"]) + " impact")

    most_impacted = df[stakeholder_col].value_counts().head(3)
    if not most_impacted.empty:
        groups = ", ".join([f"{k} ({v} changes)" for k, v in most_impacted.items()])
        summary.append(f"• Most impacted stakeholder groups: {groups}")

    neg_perceptions = df[df[perception_col].str.lower() == "negative"]
    if not neg_perceptions.empty:
        top_neg = neg_perceptions[stakeholder_col].value_counts()
        groups = ", ".join(top_neg.index[:3])
        summary.append(f"• Stakeholders with mostly negative perception: {groups}")

    high_impact = df[df[impact_col].str.lower() == "high"]
    if not high_impact.empty:
        hi_neg = high_impact[high_impact[perception_col].str.lower() == "negative"]
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
    st.title("Change Impact Analysis Summary Tool (v15.6)")
    uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])
    if uploaded_file:
        df = load_excel(uploaded_file)
        if df is not None:
            plot_impact(df)
            plot_perception(df)
            summarize_insights(df)

if __name__ == "__main__":
    main()
