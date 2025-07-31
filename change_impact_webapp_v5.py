
# Streamlit Change Impact Web App - v15.0.8
# Includes: horizontal stacked bar charts for Impact & Perception,
# proper green/yellow/red color mapping, and bar segment labels.

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.8)")
st.caption("Upload Change Impact Excel File")

uploaded_file = st.file_uploader("Drag and drop file here", type=["xlsx"], label_visibility="collapsed")
if uploaded_file:
    st.markdown(f"📄 **{uploaded_file.name}**")
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Known Change Impacts", header=2)
    except Exception as e:
        st.error(f"Failed to read worksheet: {e}")
        st.stop()

    impact_col = [col for col in df.columns if "Impact" in col][0]
    perception_col = [col for col in df.columns if "Perception" in col][0]
    stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]

    # Split comma-delimited stakeholders into individual rows
    df[stakeholder_col] = df[stakeholder_col].astype(str)
    df_expanded = df.assign(**{stakeholder_col: df[stakeholder_col].str.split(",")}).explode(stakeholder_col)
    df_expanded[stakeholder_col] = df_expanded[stakeholder_col].str.strip()

    # Define color mappings
    impact_colors = {"Low": "green", "Medium": "orange", "High": "red"}
    perception_colors = {"Positive": "green", "Neutral": "blue", "Negative": "red"}

    def plot_horizontal_stacked_chart(data, value_col, title, color_map, legend_title):
        counts = data.groupby([stakeholder_col, value_col]).size().unstack(fill_value=0)
        counts = counts.loc[:, sorted(counts.columns, key=lambda x: ["Low", "Medium", "High", "Positive", "Neutral", "Negative"].index(x))]
        counts.plot(kind="barh", stacked=True, color=[color_map.get(x, "gray") for x in counts.columns])
        plt.title(title)
        plt.xlabel("Number of Changes")
        plt.ylabel("Stakeholder Group")
        plt.legend(title=legend_title, bbox_to_anchor=(1.05, 1), loc='upper left')
        
        # Add count labels inside each bar segment
        for i, (index, row) in enumerate(counts.iterrows()):
            xpos = 0
            for cat in counts.columns:
                value = row[cat]
                if value > 0:
                    plt.text(xpos + value/2, i, str(int(value)), va='center', ha='center', fontsize=8, color="white")
                    xpos += value

        st.pyplot(plt.gcf())
        plt.clf()

    st.subheader("Change Impacts by Stakeholder Group")
    plot_horizontal_stacked_chart(df_expanded, impact_col, "Distribution of Change Impact Levels by Stakeholder", impact_colors, "Level of Impact")

    st.subheader("Perception of Change by Stakeholder")
    plot_horizontal_stacked_chart(df_expanded, perception_col, "Distribution of Change Perception Levels by Stakeholder", perception_colors, "Perception of Change")
