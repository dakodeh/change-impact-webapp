
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.6)")

def load_excel(file):
    try:
        xl = pd.ExcelFile(file)
        sheet_name = [s for s in xl.sheet_names if "Known Change Impact" in s][0]
        df = xl.parse(sheet_name, skiprows=2)
        return df
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None

def plot_impact(df):
    impact_col = [col for col in df.columns if "Impact" in col and "Level" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]

    df = df[[stakeholder_col, impact_col]].dropna()
    df[stakeholder_col] = df[stakeholder_col].astype(str)
    df[impact_col] = df[impact_col].astype(str)

    records = []
    for _, row in df.iterrows():
        stakeholders = [s.strip() for s in row[stakeholder_col].split(',')]
        for s in stakeholders:
            records.append({'Stakeholder': s, 'Impact': row[impact_col]})

    clean_df = pd.DataFrame(records)
    impact_levels = ['Low', 'Medium', 'High']
    colors = {'Low': 'green', 'Medium': 'orange', 'High': 'red'}

    impact_counts = clean_df.groupby(['Stakeholder', 'Impact']).size().unstack(fill_value=0)
    impact_counts = impact_counts.reindex(columns=impact_levels, fill_value=0)

    fig, ax = plt.subplots(figsize=(10, 6))
    impact_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[i] for i in impact_levels])
    ax.set_title("Degree of Impact by Stakeholder")
    ax.set_ylabel("Number of Changes")
    st.pyplot(fig)

def plot_perception(df):
    perception_col = [col for col in df.columns if "Perception" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]

    df = df[[stakeholder_col, perception_col]].dropna()
    df[stakeholder_col] = df[stakeholder_col].astype(str)
    df[perception_col] = df[perception_col].astype(str)

    records = []
    for _, row in df.iterrows():
        stakeholders = [s.strip() for s in row[stakeholder_col].split(',')]
        for s in stakeholders:
            records.append({'Stakeholder': s, 'Perception': row[perception_col]})

    clean_df = pd.DataFrame(records)
    perception_levels = ['Negative', 'Neutral', 'Positive']
    colors = {'Negative': 'red', 'Neutral': 'blue', 'Positive': 'green'}

    perception_counts = clean_df.groupby(['Stakeholder', 'Perception']).size().unstack(fill_value=0)
    perception_counts = perception_counts.reindex(columns=perception_levels, fill_value=0)

    fig, ax = plt.subplots(figsize=(10, 6))
    perception_counts.plot(kind="bar", stacked=True, ax=ax, color=[colors[i] for i in perception_levels])
    ax.set_title("Perception of Change by Stakeholder")
    ax.set_ylabel("Number of Changes")
    st.pyplot(fig)

uploaded_file = st.file_uploader("Upload Change Impact Analysis File", type=["xlsx"])
if uploaded_file:
    df = load_excel(uploaded_file)
    if df is not None:
        plot_impact(df)
        plot_perception(df)
