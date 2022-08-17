from cProfile import label
from cgitb import text
from collections import OrderedDict
import xlsxwriter

# Open an Excel workbook
workbook = xlsxwriter.Workbook('dict_to_excel.xlsx')

colours = {
    "person": workbook.add_format(properties={'bold': True, 'font_color': 'red'}),
    "organization": workbook.add_format(properties={'bold': True, 'font_color': 'blue'}),
    "animal": workbook.add_format(properties={'bold': True, 'font_color': 'navy'}),
    "location": workbook.add_format(properties={'bold': True, 'font_color': 'green'}),
    "time": workbook.add_format(properties={'bold': True, 'font_color': 'purple'}),
    "virus": workbook.add_format(properties={'bold': True, 'font_color': 'orange'}),
    "disease": workbook.add_format(properties={'bold': True, 'font_color': 'brown'}),
    "symptom": workbook.add_format(properties={'bold': True, 'font_color': 'cyan'}),
    "product": workbook.add_format(properties={'bold': True, 'font_color': 'magenta'}),
}

# Create a sheet
worksheet = workbook.add_worksheet('dict_data')

text_lines, label_lines = 0, 0
with open('train_text.txt') as f:
    text_lines = f.readlines()
with open('train_label_true.txt') as f:
    label_lines = f.readlines()

for i in range(len(text_lines)):
    text_line = text_lines[i].split()
    label_line = label_lines[i].split()
    worksheet.write((i+1)*2, 1, str(i+1)+".")

    for j in range(len(text_line)):
        if label_line[j][2:] in colours:
            worksheet.write((i+1)*2, j+2, text_line[j], colours[label_line[j][2:]])
        else:
            worksheet.write((i+1)*2, j+2, text_line[j])

workbook.close()