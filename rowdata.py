from xlwt import Workbook
import xlrd

read_file_name = "Data/merge.xls"

# read three sets of data from three sheets
data = xlrd.open_workbook(read_file_name)
sheet_gm = data.sheet_by_index(0)
sheet_im = data.sheet_by_index(1)
sheet_ka = data.sheet_by_index(2)

wb = Workbook()
sheet_all = wb.add_sheet('ROW')
sheet_all.write(0, 0, "ID")
sheet_all.write(0, 1, "Platform")
sheet_all.write(0, 2, "Type")
sheet_all.write(0, 3, "Response")


row=1
col=0

for user_id in range(1, 31):
    sheet_all.write(row, col, user_id)

    flag = 0
    k = 0
    m = 0
    for i in range(sheet_ka.nrows):
        col_ka = 1
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
        sheet_all.write(row, col_ka, "KA")
        sheet_all.write(row, col_ka + 1, sheet_ka.cell_value(j, 1))
        sheet_all.write(row, col_ka + 2, sheet_ka.cell_value(j, 2))
        row += 1;

    flag = 0
    k = 0
    m = 0
    for i in range(sheet_gm.nrows):
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
        sheet_all.write(row, col_gm, "GM")
        sheet_all.write(row, col_gm+2, sheet_gm.cell_value(j, 1))
        row+=1;


    flag = 0
    k = 0
    m = 0
    for i in range(sheet_im.nrows):
        col_im = 1
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
        sheet_all.write(row, col_im, "IM")
        sheet_all.write(row, col_im + 2, sheet_im.cell_value(j, 1))
        row += 1;



    row+=1

wb.save('Data/row.xls')
