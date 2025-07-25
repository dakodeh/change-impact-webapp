
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v9 - Custom Colors & Titles)")

uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])
if not uploaded_file:
    st.stop()

# Load Excel file
xls = pd.ExcelFile(uploaded_file)

# Try to identify a valid sheet
target_df = None
target_sheet = None
for sheet in xls.sheet_names:
    try:
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=1)
        df.columns = df.columns.str.strip().str.replace("\n", " ", regex=True)
        cols_lower = [col.lower() for col in df.columns]
        has_workstream = any("workstream" in col for col in cols_lower)
        has_process = any("process" in col for col in cols_lower)
        has_subprocess = any("sub-process" in col for col in cols_lower)
        has_stakeholder = any("stakeholder" in col for col in cols_lower)
        has_impact = any("impact" in col for col in cols_lower)
        has_perception = any("perception" in col for col in cols_lower)
        if (has_workstream or (has_process and has_subprocess)) and has_stakeholder and has_impact and has_perception:
            target_df = df.copy()
            target_sheet = sheet
            break
    except Exception:
        continue

if target_df is None:
    st.error("Could not find a sheet with the required Change Impact structure.")
    st.write("Sheets checked:", xls.sheet_names)
    st.stop()

st.info(f"Using sheet: **{target_sheet}**")

# Lowercase column names for detection
df = target_df
df.columns = df.columns.str.strip().str.replace("\n", " ", regex=True)
cols_lower = [col.lower() for col in df.columns]
is_new_format = any("workstream" in col for col in cols_lower)
is_old_format = any("process" in col for col in cols_lower) and any("sub-process" in col for col in cols_lower)

# Identify relevant columns
def find_column(keywords):
    for col in df.columns:
        col_lower = col.lower()
        if all(k in col_lower for k in keywords):
            return col
    return None

identifier_col = find_column(["workstream"]) if is_new_format else find_column(["process"])
stakeholder_col = find_column(["stakeholder"])
impact_col = find_column(["impact"])
perception_col = find_column(["perception"])

required = [identifier_col, stakeholder_col, impact_col, perception_col]
if any(x is None for x in required):
    st.error("Could not identify all required columns. Needed: Identifier, Stakeholder, Impact, Perception.")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

# Prepare data
df = df[[identifier_col, stakeholder_col, impact_col, perception_col]].copy()
df.columns = ["Identifier", "Stakeholder", "Impact", "Perception"]
df.dropna(subset=["Stakeholder", "Impact", "Perception"], inplace=True)
df["Stakeholder"] = df["Stakeholder"].astype(str)
df = df.assign(Stakeholder=df["Stakeholder"].str.split(",")).explode("Stakeholder")
df["Stakeholder"] = df["Stakeholder"].str.strip()

# Chart 1: Impact by Stakeholder with custom colors
impact_order = ["Low", "Medium", "High"]
impact_colors = {"Low": "green", "Medium": "orange", "High": "red"}

impact_counts = df.groupby(["Stakeholder", "Impact"]).size().unstack(fill_value=0)
impact_counts = impact_counts.reindex(columns=impact_order, fill_value=0)

st.subheader("Degree of Impact by Stakeholder")
fig1, ax1 = plt.subplots(figsize=(10, 6))
impact_counts.plot(kind="bar", stacked=True, ax=ax1, color=[impact_colors.get(col, "#333333") for col in impact_counts.columns])
ax1.set_ylabel("Number of Changes")
ax1.set_xlabel("Stakeholder Group")
ax1.set_title("Degree of Impact by Stakeholder")
ax1.legend(title="Impact Level")
st.pyplot(fig1)

# Chart 2: Perception by Stakeholder with custom colors
perception_order = ["Negative", "Neutral", "Positive"]
perception_colors = {"Negative": "red", "Neutral": "blue", "Positive": "green"}

perception_counts = df.groupby(["Stakeholder", "Perception"]).size().unstack(fill_value=0)
perception_counts = perception_counts.reindex(columns=perception_order, fill_value=0)

st.subheader("Perception of Change by Stakeholder")
fig2, ax2 = plt.subplots(figsize=(10, 6))
perception_counts.plot(kind="bar", stacked=True, ax=ax2, color=[perception_colors.get(col, "#333333") for col in perception_counts.columns])
ax2.set_ylabel("Number of Changes")
ax2.set_xlabel("Stakeholder Group")
ax2.set_title("Perception of Change by Stakeholder")
ax2.legend(title="Perception")
st.pyplot(fig2)

# Summary insight
st.subheader("Summary Insights")
high_impact_df = df[df["Impact"].str.lower() == "high"]
if high_impact_df.empty:
    st.info("No 'High' impact changes found.")
else:
    top_group = high_impact_df["Stakeholder"].value_counts().idxmax()
    count = high_impact_df["Stakeholder"].value_counts().max()
    st.markdown(f"**Interesting Fact:** The stakeholder group '{top_group}' has the highest number of high-impact changes ({count}).**")
    st.markdown(f"**Conclusion:** The visualizations highlight that **{top_group}** faces the most high-impact changes, requiring focused training or communication.")
