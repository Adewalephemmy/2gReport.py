from openpyxl import Workbook
import os

def create_excel_file(file_name):
    # Create a new workbook
    workbook = Workbook()

    # Select the active worksheet
    sheet = workbook.active

    # Set a title for the worksheet
    sheet.title = "SampleSheet"

    # Save the workbook to a file
    workbook.save(file_name)
    print(f"Excel file '{file_name}' has been created successfully!")
    print(f"Current working directory: {os.getcwd()}")
    print(f"Files in directory: {os.listdir(os.getcwd())}")

# Specify the file name
file_name = "sample_data.xlsx"
create_excel_file(file_name)
