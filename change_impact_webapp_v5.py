
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import io

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.3)")

def read_known_change_impacts_sheet(excel_file):
    xl = pd.ExcelFile(excel_file)
    for sheet in xl.sheet_names:
        if "Known Change Impacts" in sheet:
            try:
                df = xl.parse(sheet, header=2)
                return df
            except Exception as e:
                st.error(f"Failed to read sheet '{sheet}': {e}")
                return None
    st.error("No worksheet with 'Known Change Impacts' in its name was found.")
    return None

def clean_and_validate_columns(df):
    # Standardize column names
    df.columns = [str(col).strip() for col in df.columns]
    required_columns = {
        "Stakeholder Group": None,
        "Level of Impact": None,
        "Perception of Change": None
    }
    for col in df.columns:
        for key in required_columns.keys():
            if key.lower() in col.lower():
                required_columns[key] = col

    if None in required_columns.values():
        st.error("The worksheet doesn't contain recognizable Impact or Perception columns.")
        return None, None, None

    return required_columns["Stakeholder Group"], required_columns["Level of Impact"], required_columns["Perception of Change"]

def plot_impact(df, stakeholder_col, impact_col):
    st.subheader("Degree of Impact by Stakeholder")
    impact_levels = ["Low", "Medium", "High"]
    color_map = {"Low": "green", "Medium": "orange", "High": "red"}
    df_filtered = df[df[impact_col].isin(impact_levels)]
    if df_filtered.empty:
        st.warning("No recognizable impact levels found.")
        return

    impact_counts = df_filtered.groupby([stakeholder_col, impact_col]).size().unstack(fill_value=0)
    impact_counts = impact_counts[impact_levels] if all(level in impact_counts.columns for level in impact_levels) else impact_counts
    impact_counts.plot(kind="bar", stacked=True, color=[color_map.get(col, "gray") for col in impact_counts.columns])
    plt.xlabel("Stakeholder Group")
    plt.ylabel("Count of Changes")
    plt.title("Degree of Impact by Stakeholder")
    plt.legend(title="Level of Impact")
    st.pyplot(plt.gcf())
    plt.clf()

def plot_perception(df, stakeholder_col, perception_col):
    st.subheader("Perception of Change by Stakeholder")
    perception_map = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
    df_filtered = df[df[perception_col].isin(perception_map.keys())]
    if df_filtered.empty:
        st.warning("No recognizable perception values found.")
        return

    perception_counts = df_filtered.groupby([stakeholder_col, perception_col]).size().unstack(fill_value=0)
    perception_counts = perception_counts[perception_map.keys()] if all(p in perception_counts.columns for p in perception_map.keys()) else perception_counts
    perception_counts.plot(kind="bar", stacked=True, color=[perception_map.get(col, "gray") for col in perception_counts.columns])
    plt.xlabel("Stakeholder Group")
    plt.ylabel("Count of Changes")
    plt.title("Perception of Change by Stakeholder")
    plt.legend(title="Perception")
    st.pyplot(plt.gcf())
    plt.clf()

def main():
    uploaded_file = st.file_uploader("Upload Change Impact Analysis Excel File", type=["xlsx"])
    if uploaded_file:
        df = read_known_change_impacts_sheet(uploaded_file)
        if df is not None:
            stakeholder_col, impact_col, perception_col = clean_and_validate_columns(df)
            if stakeholder_col and impact_col and perception_col:
                plot_impact(df, stakeholder_col, impact_col)
                plot_perception(df, stakeholder_col, perception_col)

if __name__ == "__main__":
    main()
