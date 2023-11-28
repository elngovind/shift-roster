from openpyxl import Workbook
from openpyxl.styles import Border, Side
import calendar
import numpy as np
import pandas as pd

# Define the number of days and shifts per day
num_days = 30  # Number of days in the month

# Define the resources
resources = ['Govind', 'Pratik', 'Prithvi', 'Vikrant']

# Create a list to store the shift data
shifts_data = []

# Create a function to rotate the shifts for fairness
def rotate_shifts(shifts):
    return shifts[1:] + shifts[:1]

# Assign shifts to resources ensuring each day has all three shifts covered by different resources
shift_types = ['M', 'A', 'N']
for day in range(1, num_days + 1):
    np.random.shuffle(resources)
    shifts = shift_types.copy()
    for i, resource in enumerate(resources):
        shifts_data.append({'Date': day, 'Day_Name': calendar.day_name[(day - 1) % 7],
                            'Resource': resource, 'Shift_Type': shifts[i % len(shift_types)]})
        shifts = rotate_shifts(shifts)

# Create a DataFrame from the collected data
shifts_df = pd.DataFrame(shifts_data)

# Pivot the DataFrame to have dates as rows and resources in columns
shifts_pivot = shifts_df.pivot_table(index=['Date', 'Day_Name'], columns='Resource', values='Shift_Type', aggfunc='first')

# Create a new Workbook from openpyxl
wb = Workbook()

# Create a worksheet for the shift roster
ws = wb.active
ws.title = 'Shift Roster'

# Add headers for Dates and Days
date_column = [f"Date_{day:02d}" for day in range(1, num_days + 1)]
day_column = [calendar.day_name[(day - 1) % 7] for day in range(1, num_days + 1)]

ws.append([''] + date_column)
ws.append([''] + day_column)

# Add the shift abbreviations corresponding to resources and dates
for resource in resources:
    row = [resource]
    for day in range(1, num_days + 1):
        shift_code = shifts_pivot.loc[(day, calendar.day_name[(day - 1) % 7]), resource]
        row.append(shift_code if not pd.isnull(shift_code) else '')
    ws.append(row)

# Add border to the outside of the data
max_row = ws.max_row
max_col = ws.max_column
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
    for cell in row:
        if cell.row == 1 or cell.column == 1:
            continue
        cell.border = border

# Save the Excel file
wb.save('shift_roster_custom_format_v8.xlsx')

print("Shift roster generated and saved to 'shift_roster_custom_format_v8.xlsx'")
