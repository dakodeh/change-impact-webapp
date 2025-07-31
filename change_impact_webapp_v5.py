
# change_impact_webapp_v15_0_5.py
# This version includes horizontal bar charts and dynamic y-axis for stakeholder labels.

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.5)")

uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        if 'Known Change Impacts' in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name='Known Change Impacts', skiprows=2)
        else:
            st.warning("No 'Known Change Impacts' sheet found.")
            st.stop()
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # Clean column names
    df.columns = df.columns.str.strip()

    # Identify impact and perception columns
    impact_col = [col for col in df.columns if 'Impact' in col and 'Level' in col]
    perception_col = [col for col in df.columns if 'Perception' in col]

    # Expand multi-stakeholder rows
    if 'Stakeholder Group(s)' in df.columns:
        df['Stakeholder Group(s)'] = df['Stakeholder Group(s)'].astype(str)
        df = df.assign(**{
            'Stakeholder Group': df['Stakeholder Group(s)'].str.split(',')
        }).explode('Stakeholder Group')
        df['Stakeholder Group'] = df['Stakeholder Group'].str.strip()

    # Plot impact
    if impact_col and 'Stakeholder Group' in df.columns:
        fig, ax = plt.subplots(figsize=(10, 6))
        impact_order = ['Low', 'Medium', 'High']
        sns.countplot(
            data=df,
            y='Stakeholder Group',
            hue=impact_col[0],
            order=df['Stakeholder Group'].value_counts().index,
            hue_order=impact_order,
            palette={'Low': 'green', 'Medium': 'orange', 'High': 'red'},
            ax=ax
        )
        ax.set_title('Degree of Impact by Stakeholder')
        ax.set_xlabel('Number of Changes')
        ax.set_ylabel('Stakeholder Group')
        st.pyplot(fig)
    else:
        st.warning("The worksheet doesn't contain recognizable Impact or Stakeholder columns.")

    # Plot perception
    if perception_col and 'Stakeholder Group' in df.columns:
        fig, ax = plt.subplots(figsize=(10, 6))
        perception_order = ['Negative', 'Neutral', 'Positive']
        sns.countplot(
            data=df,
            y='Stakeholder Group',
            hue=perception_col[0],
            order=df['Stakeholder Group'].value_counts().index,
            hue_order=perception_order,
            palette={'Negative': 'red', 'Neutral': 'blue', 'Positive': 'green'},
            ax=ax
        )
        ax.set_title('Perception of Change by Stakeholder')
        ax.set_xlabel('Number of Changes')
        ax.set_ylabel('Stakeholder Group')
        st.pyplot(fig)
    else:
        st.warning("The worksheet doesn't contain recognizable Perception or Stakeholder columns.")
