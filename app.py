import pandas as pd
from openpyxl import Workbook

# Read the existing Excel file
file_path = 'drill_down_report_#81044_details.xlsx'
df = pd.read_excel(file_path)

# Extract numeric data and calculate the sum
numeric_col = df.iloc[3:19, 2].replace('[$,]', '', regex=True).apply(pd.to_numeric, errors='coerce').infer_objects()
total_sum = numeric_col.sum()

# Create a new Excel file with the sum in a single cell
workbook = Workbook()
worksheet = workbook.active

# Write the sum to the first cell (A1)
worksheet.cell(row=1, column=1, value=total_sum)

# Save the workbook
workbook.save('new_file.xlsx')
