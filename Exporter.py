import streamlit as st
import pandas as pd
import re
from io import BytesIO

# streamlit run Exporter.py

# Set title in browser and on page
st.set_page_config(page_title="Generate KUMU Input File")
st.title("Generate KUMU Input File")

# Create file uploaders
elements_file = st.file_uploader("Upload Elements CSV File", type=["csv"])
connections_file = st.file_uploader("Upload Connections CSV File", type=["csv"])

# Column names that are modified during processing
tags_col = "Tags"
filter_cols = ["From", "To"]
image_col = "Image"

# Function to extract the strings within parentheses and return the last string
def extract_parentheses(text):
    # Regex expression
    matches = re.findall(r'\((.*?)\)', str(text))
    return matches[-1] if matches else ''

# If both files have been uploaded
if elements_file and connections_file:

    # Read the CSVs
    elements_read = pd.read_csv(elements_file)
    connections_read = pd.read_csv(connections_file)

    # Replace the commas in the tags column with pipes
    if tags_col in elements_read.columns:
        elements_read[tags_col] = elements_read[tags_col].astype(str).str.replace(",", "|")
        # Remove NaN values for rows without any tags
        elements_read[tags_col] = elements_read[tags_col].replace({'nan': '', pd.NA: '', 'NaN': ''}).fillna('')
        st.success(f"Replaced commas with pipes in column: {tags_col}")
    else:
        st.warning(f"Column '{tags_col}' not found in Elements.")

    # Extract the image URL from the image column so it displays in Kumu
    if image_col in elements_read.columns:
        elements_read[image_col] = elements_read[image_col].apply(extract_parentheses)
    else:
        st.warning(f"Column '{image_col}' not found in CSV 1.")

    # Delete the first column (Name) in connections
    connections_read = connections_read.iloc[:, 1:]

    # Remove any connections that are missing a to or from value.
    missing_before = len(connections_read)
    connections_read_cleaned = connections_read.dropna(subset=filter_cols, how='any')
    missing_after = len(connections_read_cleaned)
    rows_removed = missing_before - missing_after

    # Notify how many rows were removed
    st.success(f"Removed {rows_removed} connections due to missing values in: {filter_cols}")

    # Output the marged excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        elements_read.to_excel(writer, index=False, sheet_name="Elements")
        connections_read_cleaned.to_excel(writer, index=False, sheet_name="Connections")
    output.seek(0)

    st.success("CSV files combined into one Excel with 2 sheets!")

    # Present download button
    st.download_button(
        label="Download Combined File",
        data=output,
        file_name="KUMU_Input_Sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )