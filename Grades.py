'''
team 9
section 03
This formats grades in excel
'''
import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font

inputFile = "Poorly_Organized_Data_1.xlsx"
inputWorkBook = openpyxl.load_workbook(inputFile)
inputSheet = inputWorkBook.active


myWorkbook = Workbook()
myWorkbook.remove(myWorkbook["Sheet"])


# 1. Sheets for each class
sheetNames = ["Algebra", "Calculus", "Trigonometry", "Geometry", "Statistics"]

# Loop through to add sheet if it doesn't exist already
for sheet in sheetNames:
    if sheet not in myWorkbook.sheetnames: # I got help from AI for this part (Joshua Min)
        myWorkbook.create_sheet(title=sheet)
        
# Remove default sheet
dSheet = myWorkbook.active
if dSheet.title == "Sheet":
    myWorkbook.remove(dSheet)

# 2. last name, first name, student ID, and grade columns
pass

# 3. Filter
pass

# 4. Adding functions
pass

# 5. Simple formatting
pass

# 6. Save the results
myWorkbook.save(filename = "formatted_grades.xlsx")
myWorkbook.close()


