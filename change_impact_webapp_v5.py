
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v7)")

uploaded_file = st.file_uploader("Upload a Change Impact Excel File", type=["xlsx"])
if not uploaded_file:
    st.stop()

# Load the Excel file and find the Known Change Impacts sheet
xls = pd.ExcelFile(uploaded_file)
sheet_name = "Known Change Impacts"
if sheet_name not in xls.sheet_names:
    st.error(f"Could not find the '{sheet_name}' sheet in this file.")
    st.stop()

# Read data skipping the first row, which contains helper text
df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=1)
df.columns = df.columns.str.strip().str.replace("\n", " ", regex=True)

# Lowercase columns for detection
cols_lower = [col.lower() for col in df.columns]

# Detect format
is_new_format = any("workstream" in col for col in cols_lower)
is_old_format = any("process" in col for col in cols_lower) and any("sub-process" in col for col in cols_lower)

if not is_new_format and not is_old_format:
    st.error("Unrecognized format. Expecting either 'Workstream' or 'Process/Sub-Process' in column headers.")
    st.write("Found columns:", df.columns.tolist())
    st.stop()

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

# Chart 1: Impact by Stakeholder
impact_counts = df.groupby(["Stakeholder", "Impact"]).size().unstack(fill_value=0)
st.subheader("Change Impacts by Stakeholder Group")
fig1, ax1 = plt.subplots(figsize=(10, 6))
impact_counts.plot(kind="bar", stacked=True, ax=ax1)
ax1.set_ylabel("Number of Changes")
ax1.set_xlabel("Stakeholder Group")
ax1.set_title("Change Impacts by Stakeholder")
st.pyplot(fig1)

# Chart 2: Perception by Stakeholder
perception_counts = df.groupby(["Stakeholder", "Perception"]).size().unstack(fill_value=0)
st.subheader("Change Readiness by Stakeholder Group")
fig2, ax2 = plt.subplots(figsize=(10, 6))
perception_counts.plot(kind="bar", stacked=True, ax=ax2)
ax2.set_ylabel("Number of Changes")
ax2.set_xlabel("Stakeholder Group")
ax2.set_title("Change Perception by Stakeholder")
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
