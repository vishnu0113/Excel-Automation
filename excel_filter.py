import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# Set the favicon and title for the app
st.set_page_config(page_title="Excel Automation", page_icon="📊", layout="wide")

# Function to load an Excel file and show available sheet names
def load_excel(file):
    xls = pd.ExcelFile(file)
    return xls, xls.sheet_names

# Function to clean column names
def clean_column_names(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    return df

# Function to filter data across selected sheets
def filter_sheets(xls, selected_sheets, filter_value, selected_columns):
    filtered_data = {}
    for sheet in selected_sheets:
        df = xls.parse(sheet, header=1)
        df = clean_column_names(df)
        valid_columns = [col for col in selected_columns if col in df.columns]
        if valid_columns:
            filtered_df = df[df[valid_columns].apply(lambda row: row.astype(str).str.contains(filter_value, case=False, na=False).any(), axis=1)]
            filtered_data[sheet] = filtered_df
    return filtered_data

# Function to calculate subtotals and append them to the filtered data
def calculate_subtotals(df, subtotal_columns):
    subtotal_row = {}
    for col in subtotal_columns:
        if col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                subtotal_row[col] = df[col].sum()
            else:
                subtotal_row[col] = "N/A"
    return subtotal_row

# Function to apply borders to all cells in a sheet
def apply_borders(sheet):
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = thin_border

# Function to apply bold font to the header row
def apply_bold_header(sheet):
    for cell in sheet[1]:
        cell.font = Font(bold=True)

# Function to save the filtered data and calculated subtotals into a new Excel file
def save_filtered_data(xls, selected_sheets, filtered_data, subtotal_columns):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        for sheet_name, df in filtered_data.items():
            sheet = workbook.create_sheet(sheet_name)
            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)
            
            # Calculate subtotals and add them to the bottom of the data
            subtotal_row = calculate_subtotals(df, subtotal_columns)
            sheet.append([subtotal_row.get(col, "") for col in df.columns])
            
            apply_borders(sheet)
            apply_bold_header(sheet)
    return output.getvalue()

# Streamlit UI
st.title("📊 Excel Data Filter Automation")

uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])
if uploaded_file is not None:
    xls, sheet_names = load_excel(uploaded_file)
    st.success(f"✅ File loaded successfully! Found {len(sheet_names)} sheets.")
    
    # Declare these variables in session state to retain values across tabs
    if 'filtered_data' not in st.session_state:
        st.session_state['filtered_data'] = {}
    if 'selected_columns' not in st.session_state:
        st.session_state['selected_columns'] = []
    if 'subtotal_columns' not in st.session_state:
        st.session_state['subtotal_columns'] = []

    st.subheader("🔎 Filter Data")

    # Use columns layout for better UX
    col1, col2 = st.columns(2)
    with col1:
        selected_sheets = st.multiselect("📜 Select sheets to filter", options=sheet_names)
    with col2:
        if selected_sheets:
            all_columns = set()
            for sheet in selected_sheets:
                df = xls.parse(sheet, header=1)
                df = clean_column_names(df)
                all_columns.update(df.columns.tolist())

            filter_value = st.text_input("🔍 Enter value to filter", "")
            selected_columns = st.multiselect("Select columns to filter", options=list(all_columns))
            subtotal_columns = st.multiselect("Select columns for subtotal", options=list(all_columns))

    if selected_sheets:
        st.markdown("---")
        with st.expander("Show Filtered Data Preview"):
            if filter_value and selected_columns:
                filtered_data = filter_sheets(xls, selected_sheets, filter_value, selected_columns)
                if filtered_data:
                    st.session_state['filtered_data'] = filtered_data
                    st.session_state['selected_columns'] = selected_columns
                    st.session_state['subtotal_columns'] = subtotal_columns
                    st.success("✅ Data filtered successfully!")

                    # Show preview of filtered data
                    for sheet, df in filtered_data.items():
                        st.subheader(f"📋 {sheet} - Preview")
                        st.dataframe(df.head())  # Display the first few rows

                    # Save filtered data with subtotals
                    output_data = save_filtered_data(xls, selected_sheets, filtered_data, subtotal_columns)
                    st.download_button(
                        label="⬇️ Download Filtered Excel with Subtotals", 
                        data=output_data,
                        file_name="filtered_data_with_subtotals.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ No matching data found.")
