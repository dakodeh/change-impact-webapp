
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

__version__ = "15.1.0"
st.set_page_config(layout="wide")
st.title(f"Change Impact Analysis Summary Tool (v{__version__})")

def load_data(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file)
        if "Known Change Impacts" not in xls.sheet_names:
            st.error("'Known Change Impacts' worksheet not found.")
            return None
        df_raw = pd.read_excel(xls, sheet_name="Known Change Impacts", header=2)
        return df_raw
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def clean_dataframe(df):
    df.columns = df.columns.str.strip()
    df = df.dropna(how='all')
    return df

def expand_stakeholders(df, column="Stakeholder Group(s)"):
    df = df[df[column].notna()]
    df[column] = df[column].astype(str).str.split(",")
    df = df.explode(column)
    df[column] = df[column].str.strip()
    return df

def plot_horizontal_stacked_bar(df, category_col, stack_col, title, color_map):
    if df.empty:
        st.warning(f"No data available for '{title}'")
        return

    pivot_df = df.pivot_table(index=category_col, columns=stack_col, aggfunc='size', fill_value=0)
    pivot_df = pivot_df.reindex(pivot_df.sum(axis=1).sort_values().index)

    fig, ax = plt.subplots(figsize=(10, max(4, len(pivot_df)*0.5)))
    bottom = None
    for column in pivot_df.columns:
        ax.barh(pivot_df.index, pivot_df[column], label=column, left=bottom, color=color_map.get(column, None))
        bottom = pivot_df[column] if bottom is None else bottom + pivot_df[column]
        for idx, value in enumerate(pivot_df[column]):
            if value > 0:
                ax.text(bottom[idx] - value/2, idx, str(value), ha='center', va='center', fontsize=8, color='black')

    ax.set_xlabel("Number of Changes")
    ax.set_title(title)
    ax.legend(title=stack_col, bbox_to_anchor=(1.05, 1), loc='upper left')
    st.pyplot(fig)

def plot_mitigation_pie(df):
    try:
        mitigation_cols = [c for c in df.columns if c.strip().lower() in ['comms', 'training', 'hr', 'other']]
        if not mitigation_cols:
            st.warning("No data found for mitigation strategy pie chart.")
            return

        counts = dict()
        for m_col in mitigation_cols:
            cleaned = df[m_col].astype(str).str.strip().replace('', pd.NA).dropna()
            counts[m_col] = cleaned.count()

        if sum(counts.values()) == 0:
            st.warning("Mitigation columns exist but contain no data.")
            return

        fig, ax = plt.subplots()
        wedges, texts, autotexts = ax.pie(counts.values(), labels=counts.keys(), autopct='%1.1f%%',
                                          startangle=90, pctdistance=0.85, wedgeprops=dict(width=0.3),
                                          textprops={'fontsize': 8})
        for i, a in enumerate(autotexts):
            a.set_position((a.get_position()[0]*1.2, a.get_position()[1]*1.2))
        ax.set_title("Mitigation Strategy Distribution", pad=20)
        st.pyplot(fig)
    except Exception as e:
        st.warning(f"Could not generate pie chart: {e}")

def generate_summary_insights(df):
    if "Impact Level" not in df.columns or "Stakeholder Group(s)" not in df.columns:
        return "⚠️ Summary Insights unavailable due to missing required columns."

    expanded_df = expand_stakeholders(df.copy())
    summary = ""
    high_impact = expanded_df[expanded_df["Impact Level"].str.lower() == "high"]
    if not high_impact.empty:
        most_affected = high_impact["Stakeholder Group(s)"].value_counts().idxmax()
        count = high_impact["Stakeholder Group(s)"].value_counts().max()
        summary += f"Interesting Fact: The stakeholder group '{most_affected}' has the highest number of high-impact changes ({count}).\n\n"
        summary += f"Conclusion: The visualizations highlight that {most_affected} faces the most high-impact changes, requiring focused training or communication."
    else:
        summary = "No 'High' impact changes found."

    return summary

uploaded_file = st.file_uploader("Upload your Change Impact Assessment Excel file", type=["xlsx"])
if uploaded_file:
    df = load_data(uploaded_file)
    if df is not None:
        df = clean_dataframe(df)

        if "Stakeholder Group(s)" in df.columns and "Impact Level" in df.columns:
            st.subheader("Change Impacts by Stakeholder Group")
            plot_horizontal_stacked_bar(
                expand_stakeholders(df[["Stakeholder Group(s)", "Impact Level"]].copy()),
                "Stakeholder Group(s)", "Impact Level",
                "Change Impacts by Stakeholder Group",
                {"Low": "green", "Medium": "gold", "High": "red"}
            )
        if "Stakeholder Group(s)" in df.columns and "Perception" in df.columns:
            st.subheader("Perception of Change by Stakeholder")
            plot_horizontal_stacked_bar(
                expand_stakeholders(df[["Stakeholder Group(s)", "Perception"]].copy()),
                "Stakeholder Group(s)", "Perception",
                "Perception of Change by Stakeholder",
                {"Positive": "green", "Neutral": "gray", "Negative": "red"}
            )
        st.subheader("Potential Mitigation Strategies")
        plot_mitigation_pie(df)

        st.subheader("Summary Insights")
        st.markdown(generate_summary_insights(df))

        st.markdown("---")
st.markdown(f"**Tool Version**: v{__version__}")
_Last updated: 2025-08-01 14:54_", unsafe_allow_html=True)
