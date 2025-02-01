# Python-Day-58-Automating-Excel-tasks-with-openpyxl
This project explains about the Automating Excel tasks with openpyxl

pip install openpyxl
import openpyxl

# Load an existing workbook
wb = openpyxl.load_workbook("Country_Consumption_TWH.xlsx")

# List all sheet names
print(wb.sheetnames)

# Create a new workbook
wb = openpyxl.Workbook()

# Access the active sheet
sheet = wb.active

# Rename the active sheet
sheet.title = "MySheet"

# Save the workbook
wb.save("new_file.xlsx")

# Get the value of a specific cell
value = sheet["A1"].value
print(value)

# Iterate over rows and columns
for row in sheet.iter_rows(min_row=1, max_row=5, min_col=1, max_col=3):
    for cell in row:
        print(cell.value)

# Access multiple cells
cells = sheet["A1:C3"]
for row in cells:
    for cell in row:
        print(cell.value)
for row in sheet.iter_rows(values_only=True):
    print(row)
# Write data to a specific cell
sheet["A1"] = "Hello, Excel!"

# Save the workbook
wb.save("example.xlsx")
# Add data row-wise
data = [
    ["Name", "Age", "City"],
    ["Alice", 25, "New York"],
    ["Bob", 30, "Los Angeles"],
]

for row in data:
    sheet.append(row)

# Save the workbook
wb.save("example.xlsx")
from openpyxl.styles import Font, Alignment, PatternFill

# Apply bold font to a cell
sheet["A1"].font = Font(bold=True)

# Center-align text
sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")

# Apply background color
sheet["A1"].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# Save the workbook
wb.save("formatted.xlsx")
# Add a new sheet
new_sheet = wb.create_sheet(title="New_Sheet")

# Remove a sheet
wb.remove(wb["New_file"])
# Add a formula to a cell
sheet["D1"] = "=SUM(B2:B10)"

# Save the workbook
wb.save("formulas.xlsx")
from openpyxl.chart import BarChart, Reference

# Create sample data
data = [
    ["Item", "Sales"],
    ["Apples", 50],
    ["Bananas", 30],
    ["Cherries", 40],
]
for row in data:
    sheet.append(row)

# Create a bar chart
chart = BarChart()
values = Reference(sheet, min_col=2, min_row=2, max_col=2, max_row=4)
categories = Reference(sheet, min_col=1, min_row=2, max_row=4)
chart.add_data(values, titles_from_data=True)
chart.set_categories(categories)
chart.title = "Sales Data"

# Add the chart to the sheet
sheet.add_chart(chart, "E5")

# Save the workbook
wb.save("charts.xlsx")

for row in sheet.iter_rows(min_row=1, max_row=1000, values_only=True):
    print(row)                                #Dealing with Large Datasets

