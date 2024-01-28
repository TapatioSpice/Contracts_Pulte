import streamlit as st
import openpyxl
import pandas as pd
from io import BytesIO
import base64

# Specify the GitHub raw content link to the Excel file
GITHUB_EXCEL_LINK = "https://raw.githubusercontent.com/TapatioSpice/PulteContracts/main/PulteContracts1.xlsx"

# Predefined password
PASSWORD = "landscape11"

def load_data():
    try:
        # Read the Excel file directly from the GitHub raw content link
        data = pd.read_excel(GITHUB_EXCEL_LINK)
        return data
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.stop()

# Function to filter data based on community and series
def filter_data(data, community, series):
    return data[(data['Community'] == community) & (data['Series'] == series)]

# Function to export table to Excel
def export_to_excel(data, community, series):
    # Create a Pandas Excel writer using BytesIO as the file-like buffer
    excel_buffer = BytesIO()
    excel_writer = pd.ExcelWriter(excel_buffer, engine='xlsxwriter')
    data.to_excel(excel_writer, index=False, sheet_name=f"{community}_{series}")
    excel_writer.save()

    # Save the Excel writer content to a BytesIO buffer
    excel_data = excel_buffer.getvalue()

    # Create a download link for the Excel file
    excel_b64 = base64.b64encode(excel_data).decode()
    st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="{community}_{series}.xlsx">Download Excel</a>', unsafe_allow_html=True)

# Function to export table to PDF
def export_to_pdf(data, community, series):
    # Create a PDF writer using BytesIO as the file-like buffer
    pdf_buffer = BytesIO()
    pdf_writer = pd.ExcelWriter(pdf_buffer, engine='xlsxwriter')
    data.to_excel(pdf_writer, index=False, sheet_name=f"{community}_{series}")
    pdf_writer.save()

    # Save the PDF writer content to a BytesIO buffer
    pdf_data = pdf_buffer.getvalue()

    # Create a download link for the PDF file
    pdf_b64 = base64.b64encode(pdf_data).decode()
    st.markdown(f'<a href="data:application/pdf;base64,{pdf_b64}" download="{community}_{series}.pdf">Download PDF</a>', unsafe_allow_html=True)

# Function to create and display the GUI
def create_gui(data):
    st.title("Pulte Contracts App")

    # Password protection in the sidebar
    password_input = st.sidebar.text_input("Enter password:", type="password")
    entered_password = password_input.lower()  # Convert to lowercase for case-insensitive comparison

    if entered_password != PASSWORD:
        st.sidebar.warning("Incorrect password. Please enter the correct password to proceed.")
        st.stop()

    communities = data['Community'].unique()

    community_col, series_col, button_col = st.columns([2, 2, 1])

    selected_community = community_col.selectbox('Select Community:', communities)

    series_options = data[data['Community'] == selected_community]['Series'].unique()
    selected_series = series_col.selectbox('Select Series:', series_options)

    if button_col.button('Create Table'):
        try:
            if selected_community and selected_series:
                filtered_data = filter_data(data, selected_community, selected_series)
                show_table(filtered_data)

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

# Function to create the table and display it
def show_table(data):
    data = data.sort_values(by='Work Type')

    # Round the values to 2 decimal places
    data['Amount'] = data['Amount'].round(2)

    table_data = pd.pivot_table(data, values='Amount', index='Work Type', columns='Plan', aggfunc='sum', fill_value=0)
    table_data.reset_index(inplace=True)

    # Format values in the DataFrame to display with 2 decimal places
    formatted_table_data = table_data.applymap(lambda x: f"{x:.2f}" if isinstance(x, (float, int)) else x)

    st.table(formatted_table_data)

    # Export buttons
    export_excel_button = st.button("Export to Excel")
    export_pdf_button = st.button("Export to PDF")

    if export_excel_button:
        export_to_excel(formatted_table_data, data['Community'].iloc[0], data['Series'].iloc[0])

    if export_pdf_button:
        export_to_pdf(formatted_table_data, data['Community'].iloc[0], data['Series'].iloc[0])

# Footer
footer = """
---

*Created and upkept by Alejandro Escutia | Copyright © 2024*
"""

# Create and display the GUI
create_gui(load_data())

# Add the footer
st.markdown(footer)
