
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool")

uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])
if not uploaded_file:
    st.stop()

# Load the Excel file and find the Known Change Impacts sheet
xls = pd.ExcelFile(uploaded_file)
if "Known Change Impacts" not in xls.sheet_names:
    st.error("Could not find the 'Known Change Impacts' sheet in this file.")
    st.stop()

# Read data skipping the first row, which contains helper text, not true headers
df = pd.read_excel(xls, sheet_name="Known Change Impacts", skiprows=1)

# Clean column names for easy use
df.columns = df.columns.str.strip().str.replace("\n", " ", regex=True)

# Detect format
is_new_format = "Workstream" in df.columns
is_old_format = "Process" in df.columns and "Sub-Process" in df.columns

if not is_new_format and not is_old_format:
    st.error("Unrecognized format. Must contain either 'Workstream' or 'Process/Sub-Process' columns.")
    st.stop()

# Define identifiers
identifier_col = "Workstream" if is_new_format else "Process"
stakeholder_col = next((col for col in df.columns if "Stakeholder Group" in col), None)
impact_col = next((col for col in df.columns if "Level of Impact" in col), None)
perception_col = next((col for col in df.columns if "Perception" in col), None)

if not stakeholder_col or not impact_col or not perception_col:
    st.error("Missing required columns for analysis.")
    st.stop()

# Clean and filter data
df = df[[identifier_col, stakeholder_col, impact_col, perception_col]].copy()
df.columns = ["Identifier", "Stakeholder", "Impact", "Perception"]
df.dropna(subset=["Stakeholder", "Impact", "Perception"], inplace=True)

# Expand multi-stakeholder entries
df["Stakeholder"] = df["Stakeholder"].astype(str)
df = df.assign(Stakeholder=df["Stakeholder"].str.split(",")).explode("Stakeholder")
df["Stakeholder"] = df["Stakeholder"].str.strip()

# -------------------
# Chart 1: Impact Level by Stakeholder
impact_counts = df.groupby(["Stakeholder", "Impact"]).size().unstack(fill_value=0)
st.subheader("Change Impacts by Stakeholder Group")
fig1, ax1 = plt.subplots(figsize=(10, 6))
impact_counts.plot(kind="bar", stacked=True, ax=ax1)
ax1.set_ylabel("Number of Changes")
ax1.set_xlabel("Stakeholder Group")
ax1.set_title("Change Impacts by Stakeholder")
st.pyplot(fig1)

# -------------------
# Chart 2: Perception by Stakeholder
perception_counts = df.groupby(["Stakeholder", "Perception"]).size().unstack(fill_value=0)
st.subheader("Change Readiness by Stakeholder Group")
fig2, ax2 = plt.subplots(figsize=(10, 6))
perception_counts.plot(kind="bar", stacked=True, ax=ax2)
ax2.set_ylabel("Number of Changes")
ax2.set_xlabel("Stakeholder Group")
ax2.set_title("Change Perception by Stakeholder")
st.pyplot(fig2)

# -------------------
# Summary
st.subheader("Summary Insights")
high_impact_df = df[df["Impact"].str.lower() == "high"]
if high_impact_df.empty:
    st.info("No 'High' impact changes found in the data.")
else:
    top_group = high_impact_df["Stakeholder"].value_counts().idxmax()
    count = high_impact_df["Stakeholder"].value_counts().max()
    st.markdown(f"**Interesting Fact:** The stakeholder group '{top_group}' has the highest number of high-impact changes ({count}).**")
    st.markdown(f"**Conclusion:** The visualizations highlight that **{top_group}** faces the most high-impact changes, requiring focused training or communication.")
