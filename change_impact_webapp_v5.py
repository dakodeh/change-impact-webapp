
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(layout="wide")

def get_worksheet(file):
    xl = pd.ExcelFile(file)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        if df.shape[1] > 5 and ("Impact" in df.columns.tolist() or df.columns.str.contains("Impact", case=False).any()):
            return df
    return None

def extract_data(df):
    # Clean column headers
    df.columns = df.columns.str.strip()

    # Try to get impact column
    impact_col = df.columns[3]
    stakeholder_col = [col for col in df.columns if "stakeholder" in col.lower()]
    perception_col = [col for col in df.columns if "perception" in col.lower()]

    if not stakeholder_col or not perception_col:
        return None, None, None

    # Normalize values
    df = df[[impact_col, stakeholder_col[0], perception_col[0]]].dropna()
    df.columns = ["Impact", "Stakeholder", "Perception"]
    df["Impact"] = df["Impact"].astype(str).str.strip().str.title()
    df["Perception"] = df["Perception"].astype(str).str.strip().str.title()
    df["Stakeholder"] = df["Stakeholder"].astype(str).str.strip().str.title()

    return df, "Impact", "Perception"

def plot_impact(df):
    impact_levels = ["High", "Medium", "Low"]
    colors = {"High": "red", "Medium": "orange", "Low": "green"}
    df_counts = df.groupby(["Stakeholder", "Impact"]).size().unstack(fill_value=0)
    df_counts = df_counts[[lvl for lvl in impact_levels if lvl in df_counts.columns]]

    fig, ax = plt.subplots(figsize=(10, 6))
    df_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[lvl] for lvl in df_counts.columns])
    ax.set_title("Degree of Impact by Stakeholder")
    ax.set_ylabel("Number of Changes")
    ax.set_xlabel("Stakeholder Group")
    ax.legend(title="Impact Level")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def plot_perception(df):
    perception_levels = ["Negative", "Neutral", "Positive"]
    colors = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
    df_counts = df.groupby(["Stakeholder", "Perception"]).size().unstack(fill_value=0)
    df_counts = df_counts[[lvl for lvl in perception_levels if lvl in df_counts.columns]]

    fig, ax = plt.subplots(figsize=(10, 6))
    df_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[lvl] for lvl in df_counts.columns])
    ax.set_title("Perception of Change by Stakeholder")
    ax.set_ylabel("Number of Changes")
    ax.set_xlabel("Stakeholder Group")
    ax.legend(title="Perception")
    plt.xticks(rotation=45, ha="right")
    st.pyplot(fig)

def generate_summary(df):
    total_changes = len(df)
    impact_counts = df["Impact"].value_counts().to_dict()
    top_stakeholders = df["Stakeholder"].value_counts().head(2).to_dict()
    neg_perception = df[df["Perception"] == "Negative"]["Stakeholder"].value_counts()
    high_impact_neg = df[(df["Impact"] == "High") & (df["Perception"] == "Negative")]["Stakeholder"].value_counts()
    high_impact_by_stakeholder = df[df["Impact"] == "High"]["Stakeholder"].value_counts()

    summary_lines = [
        f"• {total_changes} changes analyzed: " +
        ", ".join([f"{v} {k}" for k, v in impact_counts.items()]) + " impact",
        f"• Most impacted stakeholder groups: " +
        ", ".join([f"{k} ({v} changes)" for k, v in top_stakeholders.items()])
    ]
    if not neg_perception.empty:
        summary_lines.append("• Stakeholders with mostly negative perception: " + ", ".join(neg_perception.index))
    if not high_impact_neg.empty:
        summary_lines.append("• High impact changes perceived negatively for: " + ", ".join(high_impact_neg.index))
    if not high_impact_by_stakeholder.empty:
        summary_lines.append("• Stakeholders with multiple high impact changes: " +
                             ", ".join([f"{k} ({v})" for k, v in high_impact_by_stakeholder.items() if v > 1]))

    return "\n".join(summary_lines)

st.title("Change Impact Analysis Viewer (v14)")

uploaded_file = st.file_uploader("Upload Change Impact Excel", type=["xlsx"])
if uploaded_file:
    df_raw = get_worksheet(uploaded_file)
    if df_raw is None:
        st.error("Could not find a suitable worksheet.")
    else:
        df, impact_col, perception_col = extract_data(df_raw)
        if df is not None:
            st.subheader("Visualizations")
            plot_impact(df)
            plot_perception(df)

            st.subheader("Summary Insights")
            st.text(generate_summary(df))
        else:
            st.warning("The worksheet doesn't contain recognizable Impact or Perception columns.")
