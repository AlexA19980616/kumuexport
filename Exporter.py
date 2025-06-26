import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

# streamlit run Exporter.py

# Set title in browser and on page
st.set_page_config(page_title="Generate KUMU Input File")
st.title("Generate KUMU Input File")

st.markdown("#### Step 1: Update these column names if they are different from the defaults")

# User entered column variable names
with st.expander("⚙️ Column name settings (click to expand)", expanded=False):
    tags_col = st.text_input("Tags column name", value="Tags")
    image_col = st.text_input("Image column name", value="Bio Image")
    label_col = st.text_input("Label column name", value="Label")

    from_col = st.text_input("Connections 'From' column name", value="From")
    to_col = st.text_input("Connections 'To' column name", value="To")
    filter_cols = [from_col, to_col]

    elem_cols_remove_tf = st.text_input("Element columns to remove (comma-separated)", value="" \
    "Attachments, For Discussion, Connection (From), Connection (To), Count From Connections, Count To Connections")
    elem_cols_remove = [col.strip() for col in elem_cols_remove_tf.split(",") if col.strip()]

    conn_cols_remove_tf = st.text_input("Connection columns to remove (comma-separated)", value="Connection Title")
    conn_cols_remove = [col.strip() for col in conn_cols_remove_tf.split(",") if col.strip()]

st.markdown("#### Step 2: Update CSV files")

# Create file uploaders
elements_file = st.file_uploader("Upload Elements CSV File", type=["csv"])
connections_file = st.file_uploader("Upload Connections CSV File", type=["csv"])

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

    # Replace the commas in the tags column of Elements with pipes
    if tags_col in elements_read.columns:
        elements_read[tags_col] = elements_read[tags_col].astype(str).str.replace(",", "|")
        # Remove NaN values for rows without any tags
        elements_read[tags_col] = elements_read[tags_col].replace({'nan': '', pd.NA: '', 'NaN': ''}).fillna('')
        st.success(f"Replaced commas with pipes in column: {tags_col} in Elements")
    else:
        st.warning(f"Column '{tags_col}' not found in Elements.")

    # Extract the image URL from the image column so it displays in Kumu
    if image_col in elements_read.columns:
        elements_read[image_col] = elements_read[image_col].apply(extract_parentheses)
    else:
        st.warning(f"Column '{image_col}' not found in CSV 1.")

    # Rename the image column to image
    elements_read = elements_read.rename(columns={image_col: "Image"})

    # Remove additional columns from elements if they exist
    removed_cols = []
    for col in elem_cols_remove:
        if col in elements_read.columns:
            elements_read = elements_read.drop(columns=col)
            removed_cols.append(col)

    if removed_cols:
        st.success(f"Removed columns from Elements: {', '.join(removed_cols)}")
    else:
        if elem_cols_remove:
            st.warning("No matching columns found to remove in Elements")

    # Add a new empty 'Label' column to connections
    connections_read["Label"] = ""

    # Remove additional columns from connections if they exist
    removed_cols = []
    for col in conn_cols_remove:
        if col in connections_read.columns:
            connections_read = connections_read.drop(columns=col)
            removed_cols.append(col)

    if removed_cols:
        st.success(f"Removed columns from Connections: {', '.join(removed_cols)}")
    else:
        if conn_cols_remove:
            st.warning("No matching columns found to remove in Connections.")

    # Remove any elements that are missing a label.
    missing_before = len(elements_read)
    elements_read_cleaned = elements_read.dropna(subset=label_col, how='any')
    missing_after = len(elements_read_cleaned)
    rows_removed = missing_before - missing_after

    # Notify how many rows were removed
    st.success(f"Removed {rows_removed} elements due to missing values in: {label_col}")

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
        elements_read_cleaned.to_excel(writer, index=False, sheet_name="Elements")
        connections_read_cleaned.to_excel(writer, index=False, sheet_name="Connections")
    output.seek(0)

    # Generate dated filename
    today_str = datetime.now(ZoneInfo("Australia/Sydney")).strftime("%Y_%m_%d")
    file_name = f"KUMU_Input_Sheet_{today_str}.xlsx"

    # Present download button
    st.download_button(
        label="Download Combined File",
        data=output,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
