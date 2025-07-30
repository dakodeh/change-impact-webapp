
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.1)")

def detect_columns(df):
    col_map = {}
    for col in df.columns:
        col_l = col.lower()
        if "stakeholder" in col_l and "group" in col_l:
            col_map["stakeholder"] = col
        elif "perception" in col_l:
            col_map["perception"] = col
        elif "impact" in col_l and "level" in col_l:
            col_map["impact"] = col
    return col_map if len(col_map) == 3 else None

def plot_impact(df, impact_col, stakeholder_col):
    df = df[[stakeholder_col, impact_col]].dropna()
    df[impact_col] = df[impact_col].str.strip().str.title()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.strip()
    impact_levels = ["High", "Medium", "Low"]
    colors = {"High": "red", "Medium": "orange", "Low": "green"}
    df = df[df[impact_col].isin(impact_levels)]
    if df.empty:
        st.write("No recognizable impact levels found.")
        return
    impact_counts = df.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0)
    impact_counts = impact_counts.reindex(columns=impact_levels, fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 6))
    impact_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors.get(lvl, "gray") for lvl in impact_counts.columns])
    ax.set_title("Degree of Impact by Stakeholder")
    ax.set_ylabel("Count of Changes")
    ax.legend(title="Level of Impact")
    st.pyplot(fig)

def plot_perception(df, perception_col, stakeholder_col):
    df = df[[stakeholder_col, perception_col]].dropna()
    df[perception_col] = df[perception_col].str.strip().str.title()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.strip()
    perception_types = ["Negative", "Neutral", "Positive"]
    colors = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
    df = df[df[perception_col].isin(perception_types)]
    if df.empty:
        st.write("No recognizable perception types found.")
        return
    perception_counts = df.groupby([stakeholder_col, perception_col]).size().unstack(fill_value=0)
    perception_counts = perception_counts.reindex(columns=perception_types, fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 6))
    perception_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors.get(p, "gray") for p in perception_counts.columns])
    ax.set_title("Perception of Change by Stakeholder")
    ax.set_ylabel("Count of Changes")
    ax.legend(title="Perception of Change")
    st.pyplot(fig)

def main():
    uploaded_file = st.file_uploader("Upload your Change Impact Analysis Excel file", type=["xlsx"])
    if uploaded_file:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet = [s for s in xl.sheet_names if "impact" in s.lower()][0]
            df = xl.parse(sheet)
            col_map = detect_columns(df)
            if not col_map:
                st.error("The worksheet doesn't contain recognizable Impact or Perception columns.")
                return
            plot_impact(df, col_map["impact"], col_map["stakeholder"])
            plot_perception(df, col_map["perception"], col_map["stakeholder"])
            st.subheader("Summary Insights")
            high_impact_df = df[df[col_map["impact"]].str.lower().str.strip() == "high"]
            if not high_impact_df.empty:
                top_group = high_impact_df[col_map["stakeholder"]].mode()[0]
                st.write(f"**Interesting Fact:** The stakeholder group '{top_group}' has the highest number of high-impact changes ({high_impact_df[col_map['stakeholder']].value_counts().max()}).**")
                st.write(f"**Conclusion:** The visualizations highlight that {top_group} faces the most high-impact changes, requiring focused training or communication.")
            else:
                st.write("No 'High' impact changes found.")
        except Exception as e:
            st.error(f"Error processing file: {e}")

main()
