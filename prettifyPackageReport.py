import csv
import os
import textwrap
import datetime
import openpyxl
from openpyxl.styles import Font
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename


def format_date(date_string):
    try:
        iso_date_string = datetime.datetime.strptime(date_string, '%d/%m/%Y').strftime('%Y-%m-%d')

    except ValueError:
        iso_date_string = datetime.datetime.strptime(date_string, '%Y-%m-%d').strftime('%Y-%m-%d')

    return iso_date_string


# print instructions

print(textwrap.fill('Prettify Package Report'))
print()
print(textwrap.fill('This program will format a DM change log to be sent to the community.'))
print()
print(textwrap.fill('You will need to run a DM change log report with the following options:'))
print(textwrap.fill(' - The correct dates'))
print(textwrap.fill(' - ALL (Real Users)'))
print(textwrap.fill(' - Packages'))
print(textwrap.fill(' - New Items'))
print()
print(textwrap.fill('Click this window then press enter or return to open your file. You can press Ctrl-C at any time'
                    ' to quit.'))
input()

# open file
Tk().withdraw()
original_filename = askopenfilename(initialdir="C:\\")

# # Use for testing
# os.chdir(r'C:\Users\gemma.wright\Jisc\OneDrive - Jisc\To -do\DM change log')
# original_filename = 'original.csv'


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

    start_date_str = header_list[1][0].strip().replace(r'"', '')
    end_date_str = header_list[1][1].strip().replace(r'"', '')


    start_date = format_date(start_date_str)
    end_date = format_date(end_date_str)


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




# save file

print(textwrap.fill('The change log is ready to save. Click in this window then press return to save the file.'))
input()

outputFilename = asksaveasfilename(initialfile='New packages ' + start_date + ' to ' + end_date,
                                   defaultextension=".xlsx")

# # Use for testing
# outputFilename = 'output' + datetime.datetime.now().strftime('%Y-%m-%d %H%M%S')  + '.xlsx'

workbook.save(outputFilename)