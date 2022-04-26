import os, sys
from openpyxl import Workbook

# Get the location of execution file:
basePath = os.path.abspath(os.path.dirname(sys.argv[0]))
# Generate the path for the excel file we want to create:
exported_file = os.path.join(basePath, 'Create an excel file.xlsx')

# Create empty workbook
wb = Workbook()
# Set the variable ws as the current activated sheet
ws =  wb.active
# Create a loop that run 10 times
for x in range(1,10):
    # Assign the text "Row: n" to 10 first rows:
    ws.cell(row=x, column=2).value = 'Row: ' + str(x)
    # To merge the string "Row: " and the int x, we must convert int to string by str function.

# Try/Except is use when the line may not have error.
# For example, if you try to overwrite an openned file, you will get an error:
# [Errno 13] Permission denied
# Using Try/Except will allow you to run until the end of the script.
# It work similar to IFERROR in Excel.
try:
    wb.save(exported_file)
except Exception as e:
    print('Failed to save the result: ' + str(e))

if (os.path.isfile(exported_file)):
    print('File is created at: ', exported_file)
