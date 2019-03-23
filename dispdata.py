import xlrd

file_name = "Data/merge.xls"

# read three sets of data from three sheets
data = xlrd.open_workbook(file_name)
sheet_gm = data.sheet_by_index(0)
sheet_im = data.sheet_by_index(1)
sheet_ka = data.sheet_by_index(2)

# Take User Input
user_id = input("Enter User ID: ")
data_type = input("Enter data type (GM/IM/KA): ")
print("")


if data_type == "GM":
    flag = 0
    k = 0
    m = 0
    for i in range(sheet_gm.nrows):
        if str(sheet_gm.cell_value(i, 0)) == str(user_id)+".0":
            k = i
            flag = 1
            continue
        if flag == 1 and str(sheet_gm.cell_value(i, 0)) != '':
            m = i
            flag = 0
    if flag == 1:
        m = sheet_gm.nrows+1
    for j in range(k, m-1):
        print(str(sheet_gm.cell_value(j, 1)))


if data_type == "IM":
    flag = 0
    k = 0
    m = 0
    for i in range(sheet_im.nrows):
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
        print(str(sheet_im.cell_value(j, 1)))

if data_type == "KA":
    flag = 0
    k = 0
    m = 0
    for i in range(sheet_ka.nrows):
        if str(sheet_ka.cell_value(i, 0)) == str(user_id)+".0":
            k = i
            flag = 1
            continue
        if flag == 1 and str(sheet_ka.cell_value(i, 0)) != '':
            m = i
            flag = 0
    if flag == 1:
        m = sheet_ka.nrows+1
    for j in range(k, m-1):
        print(str(sheet_ka.cell_value(j, 2)))
