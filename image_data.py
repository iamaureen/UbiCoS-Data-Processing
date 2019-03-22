import ast
from xlwt import Workbook

with open('Data/image_data.txt') as file:
    s= file.read()
file.close()

data1 = s.replace("}", "}\n")
data2 = data1.split("\n")

wb = Workbook()
sheet1 = wb.add_sheet('Sheet1')
row = 0
col = 0

for i in data2:
    try:
        k = ast.literal_eval(i)
        for key, value in k.items():
            print(key, value)
            sheet1.write(row, col, key)
            col += 1
            for j in value:
                sheet1.write(row, col, j)
                row += 1
            col -= 1
            row += 1
    except SyntaxError:
        print("")

wb.save('Data/image.xls')







