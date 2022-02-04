from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter



workbook = load_workbook("hello_world.xlsx")
sheet = workbook['Sheet']
GameMap = workbook['GameMap']
RoomData = workbook['RoomData']
library=workbook.active


x=0
y=0

running = True


#sheet.append(["ID",'Room','Room Description', 'Paths'])
print(library.cell(1,1))

workbook.create_sheet('PlayerInfo')

library.cell(1,1)
mapx=0
mapy=0

sdsdsd
row_number=0
column_number=0
for row in sheet:
    row_number+=1
    for column in row:
        column_number+=1
        print(sheet.cell(row_number,column_number).value)

def look(roomx,roomy):

    print(sheet.cell(roomx,roomy).value)

look(3,1)


for worksheet in workbook:
    print(worksheet)




while running:

    a = input('what will you do')

    if a == 'look':
        look(3,1)




#workbook.save("hello_world.xlsx")