from json import load
from re import X
from openpyxl import load_workbook
from openpyxl.styles import Alignment



book = load_workbook('silence_labels.xlsx')
#sheet  = book.active

sheet_cells=[]
"""
for row in sheet.iter_rows():
    row_cells = []
    for cell in row:
        row_cells.append(cell.value)
    sheet_cells.append(tuple(row_cells))
#print(sheet_cells)
"""
#print(sheet_cells)

book.create_sheet('uj_tabla')
sh1 = book['uj_tabla']
sh1 = book.active


temp = []
for i in range(1,7):
    test = sh1.cell(row=1,column=i)
    temp.append(test.value)


for i in range(0,len(temp)):
    str(temp[0])
    str(temp[3])
    temp[0] = temp[0].upper()
    

sh1.merge_cells('A1:G1')  
text = temp[0]+" ("+temp[1]+")\n"+temp[2]+","+str(temp[3])+"\n"+temp[5]+","+temp[4]
targetCell = sh1.cell(row=1,column=1)
targetCell.value = text

targetCell.alignment = Alignment(horizontal='center', vertical='center')  

sh1.append
#print(text)
 #print(type(temp))

#sh1.merge_cells('A1:G1')  


#for item in sheet_cells:
 #   sh1.append(item)


book.save('silence_labels.xlsx')
print(" lefutott rendesen")