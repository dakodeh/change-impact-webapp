
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")

st.title("Change Impact Analysis Viewer")

uploaded_file = st.file_uploader("Upload a Change Impact Excel file", type=["xlsx"])

def load_data(file):
    xl = pd.ExcelFile(file)
    sheet = xl.sheet_names[0]
    df = xl.parse(sheet)
    return df

def standardize_columns(df):
    cols = [str(c).strip().lower() for c in df.columns]
    df.columns = cols
    return df

def plot_impact(df):
    impact_col = next((col for col in df.columns if 'level of impact' in col), None)
    stakeholder_col = next((col for col in df.columns if 'stakeholder group' in col), None)
    if not impact_col or not stakeholder_col:
        st.warning("No recognizable Impact or Stakeholder columns found.")
        return

    df = df[[stakeholder_col, impact_col]].dropna()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.split(',')
    df = df.explode(stakeholder_col)
    df[stakeholder_col] = df[stakeholder_col].str.strip()
    df[impact_col] = df[impact_col].str.capitalize()

    impact_levels = ['Low', 'Medium', 'High']
    df = df[df[impact_col].isin(impact_levels)]

    colors = {'Low': 'green', 'Medium': 'orange', 'High': 'red'}
    df_counts = df.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0).reindex(columns=impact_levels, fill_value=0)

    st.subheader("Degree of Impact by Stakeholder")
    fig, ax = plt.subplots(figsize=(12, 8))
    df_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[i] for i in df_counts.columns])
    plt.title("Degree of Impact by Stakeholder")
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha='right')
    plt.legend(title="Impact Level")
    st.pyplot(fig)

def plot_perception(df):
    perception_col = next((col for col in df.columns if 'perception of change' in col), None)
    stakeholder_col = next((col for col in df.columns if 'stakeholder group' in col), None)
    if not perception_col or not stakeholder_col:
        st.warning("No recognizable Perception or Stakeholder columns found.")
        return

    df = df[[stakeholder_col, perception_col]].dropna()
    df[stakeholder_col] = df[stakeholder_col].astype(str).str.split(',')
    df = df.explode(stakeholder_col)
    df[stakeholder_col] = df[stakeholder_col].str.strip()
    df[perception_col] = df[perception_col].str.capitalize()

    perception_levels = ['Negative', 'Neutral', 'Positive']
    df = df[df[perception_col].isin(perception_levels)]

    colors = {'Negative': 'red', 'Neutral': 'blue', 'Positive': 'green'}
    df_counts = df.groupby([stakeholder_col, perception_col]).size().unstack(fill_value=0).reindex(columns=perception_levels, fill_value=0)

    st.subheader("Perception of Change by Stakeholder")
    fig, ax = plt.subplots(figsize=(12, 8))
    df_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[i] for i in df_counts.columns])
    plt.title("Perception of Change by Stakeholder")
    plt.ylabel("Number of Changes")
    plt.xlabel("Stakeholder Group")
    plt.xticks(rotation=45, ha='right')
    plt.legend(title="Perception")
    st.pyplot(fig)

if uploaded_file:
    df = load_data(uploaded_file)
    df = standardize_columns(df)
    plot_impact(df)
    plot_perception(df)
