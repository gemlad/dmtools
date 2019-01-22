import csv
import os
import openpyxl
import send2trash
import textwrap
from datetime import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from openpyxl.styles import Font


# todo print instructions


# open file
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


print(package_list)

# todo put dates on excel file

workbook = openpyxl.Workbook()
workbook['Sheet'].title = 'New packages'
sheet = workbook.active



# todo put list of packages in excel file
sheet['A5'] = 'Packages added:'
for i in package_list:
    sheet.append([i])

# todo prettify


# todo save file

# Use for testing
outputFilename = 'output' + datetime.now().strftime('%Y-%m-%d %H%M%S')  + '.xlsx'

workbook.save(outputFilename)