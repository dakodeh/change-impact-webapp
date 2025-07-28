
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
def generate_summary_insights(df):
    import pandas as pd
    from collections import Counter

    if df.empty or 'Level of Impact' not in df.columns or 'Stakeholder Group(s)' not in df.columns:
        return "No summary available."

    total_changes = len(df)
    impact_counts = df['Level of Impact'].str.strip().str.capitalize().value_counts()
    impact_summary = ", ".join(f"{v} {k}" for k, v in impact_counts.items())

    # Flatten stakeholder group lists
    stakeholder_series = df['Stakeholder Group(s)'].dropna().apply(lambda x: [s.strip() for s in str(x).split(",")])
    all_stakeholders = [s for sublist in stakeholder_series for s in sublist]
    stakeholder_counts = Counter(all_stakeholders)
    top_impacted = stakeholder_counts.most_common(3)

    # Perception summary
    perception_df = df[['Level of Impact', 'Perception of Change', 'Stakeholder Group(s)']].dropna()
    perception_df['Stakeholder Group(s)'] = perception_df['Stakeholder Group(s)'].apply(lambda x: [s.strip() for s in str(x).split(",")])
    negative_stakeholders = Counter()
    high_impact_negative = set()
    high_impact_stakeholders = Counter()

    for _, row in perception_df.iterrows():
        if row['Perception of Change'].strip().lower() == "negative":
            for group in row['Stakeholder Group(s)']:
                negative_stakeholders[group] += 1
                if row['Level of Impact'].strip().lower() == "high":
                    high_impact_negative.add(group)
        if row['Level of Impact'].strip().lower() == "high":
            for group in row['Stakeholder Group(s)']:
                high_impact_stakeholders[group] += 1

    # Compose summary
    summary_lines = [
        f"📊 Insights Summary",
        f"• {total_changes} changes analyzed: {impact_summary} impact",
    ]
    if top_impacted:
        top_impacted_line = "• Most impacted stakeholder groups: " + ", ".join(f"{g} ({c} changes)" for g, c in top_impacted)
        summary_lines.append(top_impacted_line)
    if negative_stakeholders:
        summary_lines.append("• Stakeholders with mostly negative perception: " + ", ".join(negative_stakeholders.keys()))
    if high_impact_negative:
        summary_lines.append("• High impact changes perceived negatively for: " + ", ".join(high_impact_negative))
    if any(v >= 2 for v in high_impact_stakeholders.values()):
        multi_high = [k for k, v in high_impact_stakeholders.items() if v >= 2]
        summary_lines.append("• Stakeholders with multiple High impact changes: " + ", ".join(multi_high))

    return "\n".join(summary_lines)


st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.4)")

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
        has_stakeholder = any("stakeholder" in col for col in cols_lower)
        has_impact = any("impact" in col for col in cols_lower)
        has_perception = any("perception" in col for col in cols_lower)
        if (has_workstream or has_process) and has_stakeholder and has_impact and has_perception:
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

# Normalize column headers
df = target_df
df.columns = df.columns.str.strip().str.replace("\n", " ", regex=True)
cols_lower = [col.lower() for col in df.columns]

# Determine format
is_new_format = any("workstream" in col for col in cols_lower)
is_old_format = any("process" in col for col in cols_lower)

# Identify key columns
def find_column(keywords):
    for col in df.columns:
        col_lower = col.lower()
        if all(k in col_lower for k in keywords):
            return col
    return None

identifier_col = find_column(["workstream"]) if is_new_format else find_column(["process"])
stakeholder_col = find_column(["stakeholder"])
perception_col = find_column(["perception"])
impact_col = df.columns[3]  # Always use 4th column for Level of Impact

required = [identifier_col, stakeholder_col, impact_col, perception_col]
if any(x is None for x in required):
    st.error("Could not identify all required columns. Needed: Identifier, Stakeholder, Impact, Perception.")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

# Prepare data
df = df[[identifier_col, stakeholder_col, impact_col, perception_col]].copy()
df.columns = ["Identifier", "Stakeholder", "Impact", "Perception"]
df.dropna(subset=["Stakeholder", "Impact", "Perception"], inplace=True)

# Clean values
df["Impact"] = df["Impact"].astype(str).str.strip().str.title()
df["Perception"] = df["Perception"].astype(str).str.strip().str.title()
df["Stakeholder"] = df["Stakeholder"].astype(str)
df = df.assign(Stakeholder=df["Stakeholder"].str.split(",")).explode("Stakeholder")
df["Stakeholder"] = df["Stakeholder"].str.strip()

# Chart 1: Degree of Impact by Stakeholder
impact_colors = {"Low": "green", "Medium": "orange", "High": "red"}
impact_counts = df.groupby(["Stakeholder", "Impact"]).size().unstack(fill_value=0)

if impact_counts.empty:
    st.warning("No impact data available to display.")
else:
    present_impact_levels = [col for col in ["Low", "Medium", "High"] if col in impact_counts.columns]
    if not present_impact_levels:
        st.warning("No recognizable impact levels found.")
    else:
        colors = [impact_colors.get(col, "#333333") for col in present_impact_levels]
        st.subheader("Degree of Impact by Stakeholder")
        fig1, ax1 = plt.subplots(figsize=(10, 6))
        impact_counts[present_impact_levels].plot(kind="bar", stacked=True, ax=ax1, color=colors)
        ax1.set_ylabel("Number of Changes")
        ax1.set_xlabel("Stakeholder Group")
        ax1.set_title("Degree of Impact by Stakeholder")
        ax1.legend(title="Impact Level")
        st.pyplot(fig1)

# Chart 2: Perception of Change by Stakeholder
perception_colors = {"Negative": "red", "Neutral": "blue", "Positive": "green"}
perception_counts = df.groupby(["Stakeholder", "Perception"]).size().unstack(fill_value=0)
present_perception_levels = [col for col in ["Negative", "Neutral", "Positive"] if col in perception_counts.columns]
colors2 = [perception_colors.get(col, "#333333") for col in present_perception_levels]

st.subheader("Perception of Change by Stakeholder")
fig2, ax2 = plt.subplots(figsize=(10, 6))
perception_counts[present_perception_levels].plot(kind="bar", stacked=True, ax=ax2, color=colors2)
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

# trigger redeploy


# --- Level 2 Summary Insights Logic ---
def generate_level_2_insights(df):
    try:
        # Normalize column names
        df.columns = df.columns.str.strip()

        # Determine key column names
        impact_col = [col for col in df.columns if "Impact" in col and "Level" in col][0]
        perception_col = [col for col in df.columns if "Perception" in col][0]
        stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]
        workstream_col = "Workstream" if "Workstream" in df.columns else "Process"

        # Clean and explode stakeholders
        df[stakeholder_col] = df[stakeholder_col].astype(str)
        df = df.assign(Stakeholder=df[stakeholder_col].str.split(",")).explode("Stakeholder")
        df["Stakeholder"] = df["Stakeholder"].str.strip()

        # Total changes per stakeholder
        total_changes = df.groupby("Stakeholder").size()

        # Negative perception count
        negative_perception = (
            df[df[perception_col].str.lower() == "negative"]
            .groupby("Stakeholder")
            .size()
        )

        mostly_negative = negative_perception[
            (negative_perception / total_changes) > 0.5
        ]

        # High volume + negative perception
        high_volume_negative = mostly_negative[negative_perception > 2]

        # Clustered high impacts by workstream
        high_impact_clusters = (
            df[df[impact_col].str.lower() == "high"]
            .groupby(workstream_col)
            .size()
        )
        clustered_high_impact = high_impact_clusters[high_impact_clusters > 1]

        insights = []

        for stakeholder, count in mostly_negative.items():
            insights.append(f"• The {stakeholder} group received {total_changes[stakeholder]} changes, "
                            f"of which {count} were perceived negatively.")

        for stakeholder in high_volume_negative.index:
            insights.append(f"• {stakeholder} has high change volume with notable resistance — consider targeted support.")

        for ws, count in clustered_high_impact.items():
            insights.append(f"• The {ws} workstream includes {count} high-impact changes — clustered impact may require attention.")

        if not insights:
            return "No notable conditional patterns found."
        return "\n".join(insights)
    except Exception as e:
        return f"⚠️ Could not generate insights due to error: {e}"

# Inject into Streamlit display
insights = generate_level_2_insights(df)
st.markdown("### 🔍 Level 2 Insights Summary")
st.markdown(insights)
