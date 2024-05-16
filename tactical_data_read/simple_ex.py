
# # Working Simple

# import csv
# from openpyxl import load_workbook

# # Load the existing .xlsm workbook with data_only=True to get cell values
# existing_workbook = load_workbook('test.xlsm', data_only=True)

# # Get the first sheet name
# first_sheet_name = existing_workbook.sheetnames[0]

# # Select the first sheet from the existing workbook
# existing_sheet = existing_workbook[first_sheet_name]

# # Get the value in cell B2
# serial_num = existing_sheet['B2'].value

# # Print the value in cell B2 to the console
# print("Value in cell B2:", serial_num)

# # Create and open a CSV file for writing
# with open('output.csv', 'w', newline='') as csvfile:
#     csv_writer = csv.writer(csvfile)

#     # Write rows 8 to 10 from the first sheet of the existing workbook to the CSV file
#     for row_num in range(8, 11):
#         row_values = [serial_num]
#         for cell in existing_sheet[row_num]:
#             # If the cell contains a formula and refers to another sheet, cell.value will return the value from the other sheet
#             row_values.append(cell.value)
#         csv_writer.writerow(row_values)

# print("Output saved to 'output.csv'")