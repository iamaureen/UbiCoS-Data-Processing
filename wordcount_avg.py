# Program extracting first column
import xlrd

#data location
loc = ("Data/Data_Dec2018.xlsx")

wb = xlrd.open_workbook(loc)
sheetKA = wb.sheet_by_index(1)
sheetGM = wb.sheet_by_index(2)

KA = ""
GM = ""
for i in range(sheetKA.nrows):
    KA = KA+ str(sheetKA.cell_value(i, 1))+" "


lenKA = len(KA.split())
print("Words in KA: " + str(lenKA))
print("Total entries in KA: "+str(sheetKA.nrows))
print("Average words per entry in KA: "+str(lenKA/sheetKA.nrows))

print("")

for j in range(sheetGM.nrows):
    GM += str(sheetGM.cell_value(j, 1))+" "

lenGM = len(GM.split())
print("Words in GM: " + str(lenGM))
print("Total entries in GM: "+str(sheetGM.nrows))
print("Average words per entry in GM: "+str(lenGM/sheetGM.nrows))