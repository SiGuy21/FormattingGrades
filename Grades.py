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

headers = ["Last Name", "First Name", "Student ID", "Grade"]
subjects = []
subject_sheets = {}

# Create a new sheet for each unique class and add headers
for row in inputSheet.iter_rows(min_row=2, max_col=1, values_only=True):
    className = row[0]
    if className not in subjects:
        subjects.append(className)
        subject_sheets[className] = myWorkbook.create_sheet(className)
        subject_sheets[className].append(headers)
        print(f"Added {className} sheet to workbook")

# 2. last name, first name, student ID, and grade columns

# Iterate through the input sheet and add data to the corresponding class sheet
for row in inputSheet.iter_rows(min_row = 2, values_only=True):
    className = row[0]
    studentInfo = row[1]
    grade = row[2]
    # Split studentInfo into last name, first name, and student ID
    lastName, firstName, studentID = studentInfo.split('_')
    studentData = [lastName, firstName, studentID, grade]
    # Append the student data to the corresponding class sheet
    if className in subject_sheets:
        subject_sheets[className].append(studentData)
        print(f"Added data for {firstName} {lastName} to {className} sheet")


# 3. Filter
pass

# 4. Adding functions 
pass

# 5. Simple formatting
pass

# 6. Save the results
myWorkbook.save(filename = "formatted_grades.xlsx")
myWorkbook.close()
