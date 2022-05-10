from cgi import test
from json import load
from re import X
from traceback import print_tb
from openpyxl import load_workbook
from openpyxl.styles import Alignment



book = load_workbook('silence_labels.xlsx')
sheet  = book.active
max_column = sheet.max_column
max_row = sheet.max_row
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

#book.create_sheet('uj_tabla')
##h1 = book['uj_tabla']
#sh1 = book.active

temp = []
sheet_cells=[]
for idx, row in enumerate(sheet.iter_rows()):
    row_cells = []
    for cell in row:
        row_cells.append(cell.value)
        
   # print(row_cells)       
    sheet_cells.append(tuple(row_cells))
    text = row_cells[0]+" ("+row_cells[1]+")\n"+row_cells[2]+","+str(row_cells[3])+"\n"+row_cells[5]+","+row_cells[4]
   # print(text )cls
    sheet.merge_cells(start_row=idx+1, start_column=1, end_row=idx+1, end_column=7)
    targetCell = sheet.cell(row=idx+1,column=1)
    targetCell.value = text
    #temp.append(text)
    print(targetCell)
  #  print(text+" ||\n")

#print(temp)
    #sheet.append(text)
""""
for i in range(1,max_row+1):
    row_cells = []
    for j in range(1,max_column+1):
        row_cells.append(cell.value)
        print(row_cells)
        sheet_cells.append(tuple(row_cells))
    text = row_cells[0]+" ("+row_cells[1]+")\n"+row_cells[2]+","+str(row_cells[3])+"\n"+row_cells[5]+","+row_cells[4]
   
    #sheet.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
    targetCell = sheet.cell(row=i,column=i)
#print(text)  
    #targetCell.value = text
"""

#sheet.merge_cells('A1:G1')  


"""
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
"""
#print(text)
 #print(type(temp))

#sh1.merge_cells('A1:G1')  


#for item in sheet_cells:
 #   sh1.append(item)


book.save('silence_labels.xlsx')
print(" lefutott rendesen")