import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")
st.title("Change Impact Analysis Summary Tool (v15.0.3)")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Change Impact Analysis Excel File", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_name = [s for s in xls.sheet_names if "impact" in s.lower()][0]
        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=0)

        # Auto-detect key columns
        impact_col = [col for col in df.columns if "Impact" in col and "Level" in col][0]
        perception_col = [col for col in df.columns if "Perception" in col][0]
        stakeholder_col = [col for col in df.columns if "Stakeholder" in col][0]
        readiness_col = [col for col in df.columns if "Readiness" in col][0] if any("Readiness" in c for c in df.columns) else None
        workstream_col = [col for col in df.columns if "Workstream" in col or "Process" in col][0]

        # Clean up data
        df = df[[impact_col, perception_col, stakeholder_col, workstream_col] + ([readiness_col] if readiness_col else [])]
        df.dropna(subset=[stakeholder_col], inplace=True)

        # Normalize and split stakeholder groups
        df[stakeholder_col] = df[stakeholder_col].astype(str)
        df = df.assign(**{stakeholder_col: df[stakeholder_col].str.split(",")}).explode(stakeholder_col)
        df[stakeholder_col] = df[stakeholder_col].str.strip()

        st.subheader("Visualization 1: Degree of Impact by Stakeholder")
        impact_counts = pd.crosstab(df[stakeholder_col], df[impact_col])
        impact_counts = impact_counts[["Low", "Medium", "High"]] if all(x in impact_counts.columns for x in ["Low", "Medium", "High"]) else impact_counts
        colors = {"Low": "green", "Medium": "orange", "High": "red"}
        impact_counts.plot(kind="bar", stacked=True, color=[colors.get(col, "gray") for col in impact_counts.columns])
        st.pyplot(plt.gcf())
        plt.clf()

        st.subheader("Visualization 2: Perception of Change by Stakeholder")
        perception_counts = pd.crosstab(df[stakeholder_col], df[perception_col])
        colors = {"Positive": "green", "Neutral": "blue", "Negative": "red"}
        perception_counts.plot(kind="bar", stacked=True, color=[colors.get(col, "gray") for col in perception_counts.columns])
        st.pyplot(plt.gcf())
        plt.clf()

        st.subheader("Level 2 Insights Summary")
        summary = []

        # Insight 1: Top stakeholder group(s)
        top_groups = df[stakeholder_col].value_counts().head(3)
        for group, count in top_groups.items():
            summary.append(f"• {group} received {count} changes.")

        # Insight 2: Mostly negative perception
        negative_group = df[df[perception_col].str.lower() == "negative"]
        negative_summary = negative_group[stakeholder_col].value_counts()
        for group, count in negative_summary.items():
            total = df[df[stakeholder_col] == group].shape[0]
            if count / total >= 0.6:
                summary.append(f"• {group} had mostly negative perception ({count} of {total} changes).")

        # Insight 3: Volume of change + readiness
        if readiness_col:
            vol_readiness = df.groupby(stakeholder_col).agg({
                impact_col: "count",
                readiness_col: "mean"
            }).sort_values(by=impact_col, ascending=False)
            for index, row in vol_readiness.iterrows():
                if row[impact_col] >= 2 and row[readiness_col] <= 2:
                    summary.append(f"• {index} had {row[impact_col]} changes and low readiness score ({row[readiness_col]:.1f}).")

        # Insight 4: Workstreams with high impacts
        high_impact = df[df[impact_col].str.lower() == "high"]
        workstream_counts = high_impact[workstream_col].value_counts()
        for ws, count in workstream_counts.items():
            if count >= 2:
                summary.append(f"• Workstream '{ws}' has {count} High impact changes.")

        if summary:
            for line in summary:
                st.markdown(line)
        else:
            st.write("No Level 2 insights could be generated from the data.")

        # Debug info
        st.subheader("DEBUG INFO")
        st.write("Impact Column:", impact_col)
        st.write("Perception Column:", perception_col)
        st.write("Stakeholder Column:", stakeholder_col)
        st.write("Readiness Column:", readiness_col)
        st.write("Workstream Column:", workstream_col)
        st.write("Sample Data", df.head())

    except Exception as e:
        st.error(f"Error processing file: {e}")