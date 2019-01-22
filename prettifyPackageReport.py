import csv
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font

# todo print instructions


# todo open file
# Tk().withdraw()
# original_filename = askopenfilename(initialdir="C:\\")

# Use for testing
os.chdir(r'C:\Users\gemma.wright\Jisc\OneDrive - Jisc\To -do\DM change log')
original_filename = 'original.csv'


# for each row in file, add package to a list if it's not there already

with open(original_filename, encoding='utf-8') as original_file:
    for i in range(2):
        next(original_file)

    original_csv = csv.DictReader(original_file)
    package_list = []

    for row in original_csv:
        if row['Name'] in package_list:
            continue
        else:
            package_list.append(row['Name'])


# put dates on excel file

workbook = openpyxl.Workbook()
workbook['Sheet'].title = 'New packages'
sheet = workbook.active

with open(original_filename, encoding='utf-8') as original_file:
    original_header = csv.reader(original_file)
    header_list = list(original_header)
    start_date = header_list[1][0].strip().replace(r'"', '')
    end_date = datetime.strptime(header_list[1][1].strip().replace(r'"', ''), '%Y-%m-%d').strftime('%d/%m/%Y')

    print(header_list[1][0])
    print(header_list[1][1])
    print(start_date, end_date)

sheet['A1'] = 'New packages on KB+'
sheet.append(['Start date:', start_date])
sheet.append(['End date:', end_date])


# put list of packages in excel file

sheet['A5'] = 'Packages added:'
sheet['B5'] = len(package_list)
for i in package_list:
    sheet.append([i])


# prettify

font_bold = Font(bold=True)
font_not_bold = Font(bold=False)

sheet['A1'].font = font_bold
sheet['A2'].font = font_not_bold
sheet['B2'].font = font_not_bold
sheet['A5'].font = font_bold
sheet['B5'].font = font_bold

sheet.column_dimensions['A'].width = 85
sheet.column_dimensions['B'].width = 30




# todo save file

# Use for testing
outputFilename = 'output' + datetime.now().strftime('%Y-%m-%d %H%M%S')  + '.xlsx'

workbook.save(outputFilename)