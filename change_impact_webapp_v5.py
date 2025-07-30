
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.3)")

def load_data(uploaded_file):
    try:
        # Try to read the first sheet with header in the second row (row index 1)
        xl = pd.ExcelFile(uploaded_file)
        sheet_name = xl.sheet_names[0]
        df = pd.read_excel(xl, sheet_name=sheet_name, header=1)
        return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def plot_impact(df):
    impact_col = [col for col in df.columns if "Impact" in col and "Level" in col]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col]

    if not impact_col or not stakeholder_col:
        st.warning("The worksheet doesn't contain recognizable Impact or Stakeholder columns.")
        return

    impact_col = impact_col[0]
    stakeholder_col = stakeholder_col[0]

    df = df[[stakeholder_col, impact_col]].dropna()
    df[impact_col] = df[impact_col].str.strip().str.title()

    color_map = {"Low": "green", "Medium": "orange", "High": "red"}
    df[impact_col] = df[impact_col].map(lambda x: x if x in color_map else "Other")

    summary = df.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0)
    summary = summary[["Low", "Medium", "High"]] if set(["Low", "Medium", "High"]).issubset(summary.columns) else summary

    fig, ax = plt.subplots(figsize=(10, 5))
    summary.plot(kind="bar", stacked=True, color=[color_map.get(x, "gray") for x in summary.columns], ax=ax)
    plt.title("Degree of Impact by Stakeholder")
    plt.xlabel("Stakeholder")
    plt.ylabel("Count of Changes")
    st.pyplot(fig)

def plot_perception(df):
    perception_col = [col for col in df.columns if "Perception" in col]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col]

    if not perception_col or not stakeholder_col:
        st.warning("The worksheet doesn't contain recognizable Perception or Stakeholder columns.")
        return

    perception_col = perception_col[0]
    stakeholder_col = stakeholder_col[0]

    df = df[[stakeholder_col, perception_col]].dropna()
    df[perception_col] = df[perception_col].str.strip().str.title()

    color_map = {"Positive": "green", "Neutral": "blue", "Negative": "red"}
    df[perception_col] = df[perception_col].map(lambda x: x if x in color_map else "Other")

    summary = df.groupby([stakeholder_col, perception_col]).size().unstack(fill_value=0)
    fig, ax = plt.subplots(figsize=(10, 5))
    summary.plot(kind="bar", stacked=True, color=[color_map.get(x, "gray") for x in summary.columns], ax=ax)
    plt.title("Perception of Change by Stakeholder")
    plt.xlabel("Stakeholder")
    plt.ylabel("Count of Changes")
    st.pyplot(fig)

uploaded_file = st.file_uploader("Upload your Change Impact Excel file", type=["xlsx"])
if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        plot_impact(df)
        plot_perception(df)
