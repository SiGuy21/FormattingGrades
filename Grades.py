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
algebraSheet = myWorkbook.create_sheet("Algebra")
trigonometrySheet = myWorkbook.create_sheet("Trigonometry")
geometrySheet = myWorkbook.create_sheet("Geometry")
calculusSheet = myWorkbook.create_sheet("Calculus")

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


