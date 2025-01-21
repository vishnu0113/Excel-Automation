import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows

# Function to load an Excel file and show available sheet names
def load_excel(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
    return xls, sheet_names

# Function to clean column names (to handle variations like spaces, case differences, etc.)
def clean_column_names(df):
    # Strip spaces, convert to lowercase, and replace spaces with underscores
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    return df

# Function to filter data across selected sheets based on the user-provided value and selected columns
def filter_sheets(xls, selected_sheets, filter_value, selected_columns):
    filtered_data = {}
    for sheet in selected_sheets:
        # Load the sheet with the second row as the header (header=1)
        df = xls.parse(sheet, header=1)
        df = clean_column_names(df)
        
        # Check if the selected columns exist in the current sheet
        valid_columns = [col for col in selected_columns if col in df.columns]
        
        if valid_columns:
            # Filter the data based on the selected valid columns
            filtered_df = df[df[valid_columns].apply(lambda row: row.astype(str).str.contains(filter_value, case=False, na=False).any(), axis=1)]
            filtered_data[sheet] = filtered_df
    return filtered_data

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
    for cell in sheet[1]:  # The first row is the header
        cell.font = Font(bold=True)

# Function to save the filtered data into a new Excel file, including non-selected sheets
def save_filtered_data(xls, selected_sheets, filtered_data):
    with pd.ExcelWriter('filtered_data.xlsx', engine='openpyxl') as writer:
        workbook = writer.book
        
        # Write the filtered data for selected sheets
        for sheet_name, df in filtered_data.items():
            sheet = workbook.create_sheet(sheet_name)
            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)
            apply_borders(sheet)  # Apply borders for the filtered data
            apply_bold_header(sheet)  # Make the header row bold

        # Write the unfiltered data for non-selected sheets (leave as it is)
        for sheet_name in xls.sheet_names:
            if sheet_name not in selected_sheets:
                # Read the sheet data without any cleaning
                df = xls.parse(sheet_name, header=1)
                
                # Create the sheet in the workbook without modifying the data
                sheet = workbook.create_sheet(sheet_name)
                for row in dataframe_to_rows(df, index=False, header=True):
                    sheet.append(row)
                apply_borders(sheet)  # Apply borders for the unfiltered data

# Streamlit interface

st.title("Excel Data Filter Automation")

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Load the Excel file and get sheet names
    xls, sheet_names = load_excel(uploaded_file)
    
    st.write(f"Your file contains the following sheets: {', '.join(sheet_names)}")
    
    # Let the user select which sheets to filter
    selected_sheets = st.multiselect("Select sheets to filter", options=sheet_names)
    
    if selected_sheets:
        # Gather all unique column names across all sheets
        all_columns = set()
        for sheet in selected_sheets:
            df = xls.parse(sheet, header=1)
            df = clean_column_names(df)
            all_columns.update(df.columns.tolist())

        # Show the user the unique column names from the selected sheets
        st.write("Unique column names across selected sheets:")
        st.write(all_columns)
        
        # Ask for the filter value
        filter_value = st.text_input("Enter the filter value (e.g., Elephant)", "Elephant")
        
        if filter_value:
            # Let the user select which columns they want to filter
            selected_columns = st.multiselect("Select columns to filter", options=list(all_columns))
            
            if selected_columns:
                # Filter the data based on user input and selected columns
                filtered_data = filter_sheets(xls, selected_sheets, filter_value, selected_columns)
                
                if filtered_data:
                    # Display the filtered data preview
                    for sheet, df in filtered_data.items():
                        st.write(f"Filtered data from sheet: {sheet}")
                        st.dataframe(df)

                    # Provide download link for filtered data
                    save_filtered_data(xls, selected_sheets, filtered_data)
                    with open("filtered_data.xlsx", "rb") as file:
                        st.download_button(
                            label="Download Filtered Excel File",
                            data=file,
                            file_name="filtered_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.write("No matching data found for the filter criteria.")
