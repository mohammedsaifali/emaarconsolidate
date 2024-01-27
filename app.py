import streamlit as st
import pandas as pd
import io

# Function to process each sheet in an Excel file
def process_sheet(sheet_data):
    # Set the third row as the header
    sheet_data.columns = sheet_data.iloc[2]
    sheet_data = sheet_data[3:]
    
    # Drop rows with fewer than two non-NA values
    sheet_data.dropna(thresh=2, inplace=True)
    
    # Convert data types for 'Cl. Qty' and 'Cl. Value' to numeric, errors are coerced to NaN
    sheet_data['Cl. Qty'] = pd.to_numeric(sheet_data['Cl. Qty'], errors='coerce')
    sheet_data['Cl. Value'] = pd.to_numeric(sheet_data['Cl. Value'], errors='coerce')
    
    return sheet_data

# Streamlit UI
st.title('Emaar Stock Trend Analysis')

# File uploader allows user to add multiple files
uploaded_files = st.file_uploader("Upload Excel files", accept_multiple_files=True, type=['xls', 'xlsx'])

if uploaded_files:
    all_processed_sheets = []

    for uploaded_file in uploaded_files:
        # Read all sheets from each uploaded file
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        for sheet_name, sheet_data in all_sheets.items():
            processed_sheet = process_sheet(sheet_data)
            all_processed_sheets.append(processed_sheet)

    # Concatenate all processed sheets into a single DataFrame
    merged_data = pd.concat(all_processed_sheets, ignore_index=True)

    # Group by 'Item Code' and 'Particulars' and sum 'Cl. Qty' and 'Cl. Value'
    aggregated_data = merged_data.groupby(['Item Code', 'Particulars']).agg({
        'Cl. Qty': 'sum', 
        'Cl. Value': 'sum'
    }).reset_index()

    # Display aggregated data
    st.write('Aggregated Data:', aggregated_data)

    # Convert aggregated data to Excel and then to bytes for the download button
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        aggregated_data.to_excel(writer, index=False)
    st.download_button(label="Download Aggregated Data as Excel", data=output.getvalue(), file_name='aggregated_data.xlsx', mime='application/vnd.ms-excel')
