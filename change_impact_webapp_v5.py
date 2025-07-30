
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.3)")

def load_excel(file):
    try:
        xls = pd.ExcelFile(file)
        sheet = [s for s in xls.sheet_names if "impact" in s.lower()][0]
        df = xls.parse(sheet)
        df.columns = df.columns.str.strip()  # Trim whitespace
        return df
    except Exception as e:
        st.error(f"Error loading Excel: {e}")
        return None

def identify_columns(df):
    col_map = {}
    for col in df.columns:
        name = col.strip().lower()
        if "stakeholder" in name:
            col_map["stakeholder"] = col
        elif "perception" in name:
            col_map["perception"] = col
        elif "impact" in name and "level" in name:
            col_map["impact"] = col
        elif "workstream" in name and "sub" not in name:
            col_map["workstream"] = col
    return col_map

def plot_impact_chart(df, col_map):
    df[col_map["impact"]] = df[col_map["impact"]].str.strip().str.title()
    df[col_map["stakeholder"]] = df[col_map["stakeholder"]].astype(str).str.split(",")
    exploded_df = df.explode(col_map["stakeholder"])
    exploded_df[col_map["stakeholder"]] = exploded_df[col_map["stakeholder"]].str.strip()

    order = ["Low", "Medium", "High"]
    colors = {"Low": "green", "Medium": "orange", "High": "red"}
    crosstab = pd.crosstab(exploded_df[col_map["stakeholder"]], exploded_df[col_map["impact"]])
    crosstab = crosstab[[lvl for lvl in order if lvl in crosstab.columns]]

    fig, ax = plt.subplots(figsize=(10, 6))
    crosstab.plot(kind="bar", stacked=True, color=[colors[lvl] for lvl in crosstab.columns], ax=ax)
    ax.set_title("Degree of Impact by Stakeholder")
    ax.set_ylabel("Number of Changes")
    ax.legend(title="Level of Impact", loc="upper right")
    st.pyplot(fig)

def plot_perception_chart(df, col_map):
    df[col_map["perception"]] = df[col_map["perception"]].str.strip().str.title()
    df[col_map["stakeholder"]] = df[col_map["stakeholder"]].astype(str).str.split(",")
    exploded_df = df.explode(col_map["stakeholder"])
    exploded_df[col_map["stakeholder"]] = exploded_df[col_map["stakeholder"]].str.strip()

    colors = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
    order = ["Negative", "Neutral", "Positive"]
    crosstab = pd.crosstab(exploded_df[col_map["stakeholder"]], exploded_df[col_map["perception"]])
    crosstab = crosstab[[p for p in order if p in crosstab.columns]]

    fig, ax = plt.subplots(figsize=(10, 6))
    crosstab.plot(kind="bar", stacked=True, color=[colors[p] for p in crosstab.columns], ax=ax)
    ax.set_title("Perception of Change by Stakeholder")
    ax.set_ylabel("Number of Changes")
    ax.legend(title="Perception", loc="upper right")
    st.pyplot(fig)

def generate_summary(df, col_map):
    try:
        df[col_map["stakeholder"]] = df[col_map["stakeholder"]].astype(str).str.split(",")
        exploded_df = df.explode(col_map["stakeholder"])
        exploded_df[col_map["stakeholder"]] = exploded_df[col_map["stakeholder"]].str.strip()
        exploded_df[col_map["impact"]] = exploded_df[col_map["impact"]].str.strip().str.title()
        exploded_df[col_map["perception"]] = exploded_df[col_map["perception"]].str.strip().str.title()

        summary = "### Summary Insights\n"
        impact_counts = exploded_df[col_map["impact"]].value_counts().to_dict()
        stakeholder_counts = exploded_df[col_map["stakeholder"]].value_counts().to_dict()
        high_impact = exploded_df[exploded_df[col_map["impact"]] == "High"]
        high_by_stakeholder = high_impact[col_map["stakeholder"]].value_counts().to_dict()

        summary += f"- {sum(impact_counts.values())} changes analyzed: " + ", ".join(f"{k} ({v})" for k, v in impact_counts.items()) + "\n"

        if stakeholder_counts:
            top_stakeholders = sorted(stakeholder_counts.items(), key=lambda x: x[1], reverse=True)[:2]
            summary += "- Most impacted stakeholder groups: " + ", ".join(f"{k} ({v} changes)" for k, v in top_stakeholders) + "\n"

        perception_by_stakeholder = exploded_df.groupby(col_map["stakeholder"])[col_map["perception"]].apply(lambda x: x.value_counts().idxmax())
        mostly_negative = perception_by_stakeholder[perception_by_stakeholder == "Negative"].index.tolist()
        if mostly_negative:
            summary += "- Stakeholders with mostly negative perception: " + ", ".join(mostly_negative) + "\n"

        if high_by_stakeholder:
            repeated = [k for k, v in high_by_stakeholder.items() if v > 1]
            if repeated:
                summary += "- Stakeholders with multiple High impact changes: " + ", ".join(repeated) + "\n"

        st.markdown(summary)
    except Exception as e:
        st.error(f"Failed to generate summary insights: {e}")

uploaded = st.file_uploader("Upload your Change Impact Excel file", type=["xlsx"])
if uploaded:
    df = load_excel(uploaded)
    if df is not None:
        col_map = identify_columns(df)
        if all(k in col_map for k in ["stakeholder", "perception", "impact"]):
            plot_impact_chart(df, col_map)
            plot_perception_chart(df, col_map)
            generate_summary(df, col_map)
        else:
            st.warning("The worksheet doesn't contain recognizable Impact or Perception columns.")
