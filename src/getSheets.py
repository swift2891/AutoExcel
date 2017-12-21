import os
import openpyxl


for the_file in os.listdir('uploads'):
    file_path = os.path.join('uploads', the_file)
    try:
        if the_file.endswith('xlsx'):
            targetInputFile = file_path
    except Exception as e:
        print(e)

# Input Excel
print("Loading Input file: " + targetInputFile)
wb = openpyxl.load_workbook(targetInputFile)
# Get Correct Sheet
listOfSheets = wb.get_sheet_names()