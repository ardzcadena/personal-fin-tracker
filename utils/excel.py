import openpyxl

# This script reads an Excel file and converts it into a list of dictionaries.
def read_excel(file_path):
    try:
        # Load the workbook
        workbook = openpyxl.load_workbook(file_path, data_only=True)

        # Get the first sheet
        sheet = workbook.active

        # Read the data into a list of dictionaries
        data = []
        headers = [cell.value for cell in sheet[1]]  # Assuming the first row is the header
        for row in sheet.iter_rows(min_row=2, values_only=True):
            row_data = {headers[i]: row[i] for i in range(len(headers))}
            data.append(row_data)

        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []