import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
import os
import shutil
from io import BytesIO
import base64

# Streamlit app configuration
st.set_page_config(page_title="Change Impact Analysis", layout="wide")
st.title("Change Impact Analysis Dashboard")
st.write("Upload your 'Change Impact Analysis Template.xlsx' to generate charts and summary.")

# File uploader
uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx"])

if uploaded_file:
    # Define output directory
    output_dir = "Change_Impact_Charts"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # Save uploaded file temporarily
    input_file = "temp_uploaded_file.xlsx"
    with open(input_file, "wb") as f:
        f.write(uploaded_file.getvalue())

    # Read the Excel file
    try:
        df = pd.read_excel(input_file, sheet_name="Known Change Impacts", skiprows=[0, 2])
    except FileNotFoundError:
        st.error(f"Error: {input_file} not found.")
        st.stop()
    except ValueError as e:
        st.error(f"Error: Sheet 'Known Change Impacts' not found or invalid. Details: {e}")
        st.stop()

    # Rename columns
    column_mapping = {
        col: "Process" for col in df.columns if "Identify the impacted process" in str(col)
    }
    column_mapping.update({
        col: "Stakeholder Group(s)" for col in df.columns if "List impacted stakeholder groups" in str(col)
    })
    column_mapping.update({
        col: "Level of Impact" for col in df.columns if "H,M,L" in str(col)
    })
    column_mapping.update({
        col: "Perception of Change" for col in df.columns if "How will change be perceived" in str(col)
    })
    df = df.rename(columns=column_mapping)

    # Verify required columns
    required_columns = ["Process", "Stakeholder Group(s)", "Level of Impact", "Perception of Change"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing required columns: {missing_columns}")
        st.write("Available columns:", list(df.columns))
        st.stop()

    # Filter invalid rows
    df = df[df["#"].notna() & df["#"].apply(lambda x: str(x).isdigit())]
    df = df[required_columns].dropna(subset=["Process", "Stakeholder Group(s)", "Level of Impact"])

    # Display filtered data
    st.subheader("Filtered Data")
    st.write(df)

    if df.empty:
        st.error("No valid data found. Ensure rows 4+ contain numeric '#', non-empty 'Process', 'Stakeholder Group(s)', and 'Level of Impact'.")
        st.stop()

    # Extract unique stakeholder groups
    stakeholders = set()
    for groups in df["Stakeholder Group(s)"].dropna():
        for group in [g.strip() for g in groups.split(",") if g.strip()]:
            stakeholders.add(group.lower())
    stakeholders = sorted(stakeholders)
    if not stakeholders:
        st.error("No stakeholder groups found in the data.")
        st.stop()

    st.write("Detected stakeholder groups:", stakeholders)

    # Initialize summaries
    impact_summary = {s: {"High": 0, "Medium": 0, "Low": 0, "N/A": 0} for s in stakeholders}
    readiness_summary = {s: {"Positive": 0, "Negative": 0, "Neutral": 0, "N/A": 0} for s in stakeholders}

    # Process stakeholder groups
    for _, row in df.iterrows():
        groups = [g.strip() for g in row["Stakeholder Group(s)"].split(",") if g.strip()]
        impact = row["Level of Impact"]
        perception = row["Perception of Change"] if pd.notna(row["Perception of Change"]) else "N/A"
        for group in groups:
            group_lower = group.lower()
            if group_lower in stakeholders:
                if impact in impact_summary[group_lower]:
                    impact_summary[group_lower][impact] += 1
                if perception in readiness_summary[group_lower]:
                    readiness_summary[group_lower][perception] += 1

    st.write("Impact summary:", impact_summary)
    st.write("Readiness summary:", readiness_summary)

    if all(sum(counts.values()) == 0 for counts in impact_summary.values()):
        st.error("No valid stakeholder data processed. Ensure 'Stakeholder Group(s)' contains non-empty values and 'Level of Impact' is High, Medium, Low, or N/A.")
        st.stop()

    # Prepare data for charts
    impact_data = pd.DataFrame({
        "Stakeholder": stakeholders,
        "High": [impact_summary[s]["High"] for s in stakeholders],
        "Medium": [impact_summary[s]["Medium"] for s in stakeholders],
        "Low": [impact_summary[s]["Low"] for s in stakeholders],
        "N/A": [impact_summary[s]["N/A"] for s in stakeholders]
    })

    readiness_data = pd.DataFrame({
        "Stakeholder": stakeholders,
        "Positive": [readiness_summary[s]["Positive"] for s in stakeholders],
        "Negative": [readiness_summary[s]["Negative"] for s in stakeholders],
        "Neutral": [readiness_summary[s]["Neutral"] for s in stakeholders],
        "N/A": [readiness_summary[s]["N/A"] for s in stakeholders]
    })

    # Plot Change Impacts
    st.subheader("Change Impacts by Stakeholder Group")
    fig1, ax1 = plt.subplots(figsize=(max(10, len(stakeholders) * 0.5), 6))
    impact_data.set_index("Stakeholder")[["High", "Medium", "Low", "N/A"]].plot(
        kind="bar", stacked=True, color=["#ef4444", "#f97316", "#22c55e", "#6b7280"], ax=ax1
    )
    ax1.set_title("Change Impacts by Stakeholder Group")
    ax1.set_xlabel("Stakeholder Group")
    ax1.set_ylabel("Number of Changes")
    ax1.legend(title="Impact Level")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    st.pyplot(fig1)

    # Save and provide download for change_impacts.png
    img1_buffer = BytesIO()
    fig1.savefig(img1_buffer, format="png")
    img1_buffer.seek(0)
    st.download_button(
        label="Download Change Impacts Chart",
        data=img1_buffer,
        file_name="change_impacts.png",
        mime="image/png"
    )
    plt.close(fig1)

    # Plot Change Readiness
    st.subheader("Change Readiness by Stakeholder Group")
    fig2, ax2 = plt.subplots(figsize=(max(10, len(stakeholders) * 0.5), 6))
    readiness_data.set_index("Stakeholder")[["Positive", "Negative", "Neutral", "N/A"]].plot(
        kind="bar", color=["#22c55e", "#ef4444", "#3b82f6", "#6b7280"], ax=ax2
    )
    ax2.set_title("Change Readiness by Stakeholder Group")
    ax2.set_xlabel("Stakeholder Group")
    ax2.set_ylabel("Number of Responses")
    ax2.legend(title="Perception")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    st.pyplot(fig2)

    # Save and provide download for change_readiness.png
    img2_buffer = BytesIO()
    fig2.savefig(img2_buffer, format="png")
    img2_buffer.seek(0)
    st.download_button(
        label="Download Change Readiness Chart",
        data=img2_buffer,
        file_name="change_readiness.png",
        mime="image/png"
    )
    plt.close(fig2)

    # Generate and display summary
    max_high_impact = max(impact_summary[s]["High"] for s in stakeholders)
    high_impact_stakeholder = next(s for s in stakeholders if impact_summary[s]["High"] == max_high_impact)
    summary = f"Interesting Fact: The stakeholder group '{high_impact_stakeholder}' has the highest number of high-impact changes ({max_high_impact}).\n"
    summary += f"Conclusion: The visualizations highlight that {high_impact_stakeholder} face the most high-impact changes, requiring focused training."
    st.subheader("Summary")
    st.write(summary)

    # Save and provide download for summary.txt
    summary_buffer = BytesIO(summary.encode())
    st.download_button(
        label="Download Summary",
        data=summary_buffer,
        file_name="summary.txt",
        mime="text/plain"
    )

    # Clean up temporary file
    if os.path.exists(input_file):
        os.remove(input_file)
else:
    st.info("Please upload an Excel file to proceed.")