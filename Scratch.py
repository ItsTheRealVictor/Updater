import openpyxl as opx
import pprint
import datetime
from dateutil import parser
from dateutil import relativedelta


wb = opx.load_workbook(r"C:\Users\VD102541\Desktop\Copy of West Lab Chamber Usage WW25 2021.xlsx")

# List of sheets 
sheetsList = wb.sheetnames
chambersOnlySheetsList = sheetsList[0:14]

# print(wb[sheetsList[3]])






chamberChoice = ''
chamberList = {
    "1": "125C Bake",
    "2": "150C Bake",
    "3": "180C Bake",
    "4": "210C Bake",
    "5": "30.60 Soak",
    "6": "60.60 Soak",
    "7": "85.85 Soak",
    "8": "Oven 8 Cold",
    "9": "UHAST 1",
    "10": "UHAST 2", 
    "11": "UHAST 5",
    "12": "TC.2_D",
    "13": "TC.3_D",
    "14": "TC.4_D"
}

pprint.pprint(chamberList)
userChamberSelection = input("Enter the number which corresponds to the chamber:")

for k, v in chamberList.items():
    if userChamberSelection == k:
        chamberChoice = f"{v}"



insertLocation = 3
for index, value in enumerate(chambersOnlySheetsList):
    if value == chamberChoice:
        wb[sheetsList[index]].insert_rows(insertLocation)

wb.save('InsertRowTest.xlsx')
print('Did this work?')

