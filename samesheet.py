from xlwt import Workbook
import xlrd

read_file_name = "Data/merge.xls"

# read three sets of data from three sheets
data = xlrd.open_workbook(read_file_name)
sheet_gm = data.sheet_by_index(0)
sheet_im = data.sheet_by_index(1)
sheet_ka = data.sheet_by_index(2)

wb = Workbook()
sheet_all = wb.add_sheet('ALL')
sheet_all.write(0, 0, "ID")
sheet_all.write(0, 1, "GM")
sheet_all.write(0, 2, "IM")
sheet_all.write(0, 3, "KA")

row=1
col=0

for user_id in range(1, 31):
    sheet_all.write(row, col, user_id)
    max_row = 0
    flag = 0
    k = 0
    m = 0
    for i in range(sheet_gm.nrows):
        row_gm = row
        col_gm = 1
        if str(sheet_gm.cell_value(i, 0)) == str(user_id)+".0":
            k = i
            flag = 1
            continue
        if flag == 1 and str(sheet_gm.cell_value(i, 0)) != '':
            m = i
            flag = 0
    if flag == 1:
        m = sheet_gm.nrows + 1
    for j in range(k, m - 1):
        sheet_all.write(row_gm, col_gm, sheet_gm.cell_value(j, 1))
        row_gm+=1;
        if row_gm>max_row:
            max_row = row_gm

    flag = 0
    k = 0
    m = 0
    for i in range(sheet_im.nrows):
        row_im = row
        col_im = 2
        if str(sheet_im.cell_value(i, 0)) == str(user_id)+".0":
            k = i
            flag = 1
            continue
        if flag == 1 and str(sheet_im.cell_value(i, 0)) != '':
            m = i
            flag = 0
    if flag == 1:
        m = sheet_im.nrows+1
    for j in range(k, m-1):
        sheet_all.write(row_im, col_im, sheet_im.cell_value(j, 1))
        row_im += 1;
        if row_im > max_row:
            max_row = row_im

    flag = 0
    k = 0
    m = 0
    for i in range(sheet_ka.nrows):
        row_ka = row;
        col_ka = 3
        if str(sheet_ka.cell_value(i, 0)) == str(user_id)+".0":
            k = i
            flag = 1
            continue
        if flag == 1 and str(sheet_ka.cell_value(i, 0)) != '':
            m = i
            flag = 0
    if flag == 1:
        m = sheet_ka.nrows + 1
    for j in range(k, m - 1):
        sheet_all.write(row_ka, col_ka, sheet_ka.cell_value(j, 2))
        row_ka += 1;
        if row_ka > max_row:
            max_row = row_ka
    if row == max_row or max_row==0:
        row+=1
    else:
        row = max_row+1

wb.save('Data/all.xls')
