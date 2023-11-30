import openpyxl

excel_file = "Interest_import_test.xlsx"

try:
    workbook = openpyxl.load_workbook(excel_file, data_only=True)
except FileNotFoundError:
    print(f"Error: The file '{excel_file}' was not found.")
    exit(1)
except openpyxl.utils.exceptions.InvalidFileException:
    print(f"Error: '{excel_file}' is not a valid Excel file.")
    exit(1)
except Exception as e:
    print(f"An error occurred while loading the Excel file: {e}")
    exit(1)

sheet = workbook["interest_calc"]

data = []

for row in sheet.iter_rows(min_row=2, values_only=True):
    principal, rate, time, simple_interest, compound_interest = row[0], row[1], row[2], row[3], row[4]
    data.append((principal, rate, time, simple_interest, compound_interest))

for row_data in data:
    print("Principal Amount:", row_data[0])
    print("Interest Rate (%):", row_data[1])
    print("Time Period (Years):", row_data[2])
    print("Simple Interest:", row_data[3])
    print("Compound Interest:", row_data[4], "\n")

workbook.close()
