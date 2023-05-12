import pandas as pd
from openpyxl import load_workbook
from openpyxl import load_workbook

# Load the workbook
wb = load_workbook('DKSalaries.xlsx')

# Select the active sheet
sheet = wb['Sheet1']  # replace 'Sheet1' with the name of your sheet

# Define the columns you want to keep
columns_to_keep = ['Position', 'Name', 'ID', 'Salary', 'Game Info', 'AvgPointsPerGame']

# Get the max column count
max_column = sheet.max_column

# Traverse in reverse order (to avoid index change problem during deletion)
for i in range(max_column, 0, -1):
    # If the column header is not in columns_to_keep list, delete it
    if sheet.cell(row=1, column=i).value not in columns_to_keep:
        sheet.delete_cols(i)

# Add a new column header
sheet['G1'] = 'Predicted Points'  # Assuming 'G' is the next empty column

# Save the workbook
wb.save('DKSalaries.xlsx')

# Load the CSV file
df = pd.read_excel('DKSalaries.xlsx', sheet_name='Sheet1')  # Replace 'Sheet1' with your sheet name

# Display the result
print(df)


