import ast
from xlwt import Workbook

# Read general message text file
with open('Data/gm_data.txt') as file_gm:
    text_gm = file_gm.read()
file_gm.close()

# Read image text file
with open('Data/image_data.txt') as file_im:
    text_im = file_im.read()
file_im.close()

# Read Khan Academy text file
with open('Data/ka_data.txt', encoding='utf-8') as file_ka:
    text_ka = file_ka.read()
file_ka.close()

# Data Pre-processing
data1_gm = text_gm.replace("}", "}\n")
data2_gm = data1_gm.split("\n")

data1_im = text_im.replace("}", "}\n")
data2_im = data1_im.split("\n")

data1_ka = text_ka.replace("]}", "]}\n")
data2_ka = data1_ka.split("\n")

# Create excel file and three different sheets
wb = Workbook()
sheet_gm = wb.add_sheet('GM')
sheet_im = wb.add_sheet('IM')
sheet_ka = wb.add_sheet('KA')

sheet_ka.write(0, 0, "ID")
sheet_ka.write(0, 1, "Type")
sheet_ka.write(0, 2, "Response")

row = 0
col = 0

# parse general message data and write to excel
for i in data2_gm:
    try:
        k = ast.literal_eval(i)
        for key, value in k.items():
            sheet_gm.write(row, col, key)
            col += 1
            for j in value:
                sheet_gm.write(row, col, j)
                row += 1
            col -= 1
            row += 1
    except SyntaxError:
        print("")

row = 0
col = 0

# Parse image data and write to excel
for i in data2_im:
    try:
        k = ast.literal_eval(i)
        for key, value in k.items():
            sheet_im.write(row, col, key)
            col += 1
            for j in value:
                sheet_im.write(row, col, j)
                row += 1
            col -= 1
            row += 1
    except SyntaxError:
        print("")

row = 1
col = 0

# Parse khan academy data and write to excel
for i in data2_ka:
    try:
        k = ast.literal_eval(i)
        for key1, value1 in k.items():
            sheet_ka.write(row, col, key1)
            for j in value1:
                for key2, value2 in j.items():
                    if key2 == "type":
                        if value2 != '':
                            sheet_ka.write(row, col+1, value2)
                    if key2 == "response":
                        if value2 != '':
                            sheet_ka.write(row, col+2, value2)
                            row += 1
        row += 1
    except SyntaxError:
        print("")
wb.save('Data/merge.xls')
