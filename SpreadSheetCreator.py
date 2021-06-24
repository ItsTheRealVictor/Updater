import openpyxl as opx
import pprint

wb = opx.load_workbook(r"C:\Users\VD102541\Desktop\Copy of West Lab Chamber Usage WW25 2021.xlsx")
sheetBakeHTS = wb['Bake_HTS']

userLotNumber = input('Enter the Lot number: ')
userPartNumber = input('Enter the part number: ')
userNumOfLots = input('Enter the number of lots: ')
userQuantity = input('Enter the quantity: ')
userStartTime = input('Enter the starting time: ')
userLotOwner = input('Enter the owner: ')
chamberChoice = ''
chamberList = {
    "1": "125C Bake",
    "2": "150C Bake",
    "3": "180C Bake",
    "4": "210C Bake",
    "5": "225C Bake (Wafers Only)",
    "6": "85C/60RH Soak 1",
    "7": "60C/60RH Soak 5",
    "8": "30C/60RH Soak 2",
    "9": "UHAST 1",
    "10": "UHAST 2", 
    "11": "UHAST 4",
    "12": "UHAST 5",
    "13": "TC2",
    "14": "TC 3"
}

pprint.pprint(chamberList)
userChamberSelection = input("Enter the number which corresponds to the chamber:")

for k, v in chamberList.items():
    if userChamberSelection == k:
        chamberChoice = f"{v}"

userDataList = ['', '', 
                    userLotNumber, 
                    userPartNumber, 
                    userNumOfLots, 
                    userQuantity, 
                    parsedUserDateInput, 
                    calculateFutureDate(userDateInput), 
                    userStartTime, 
                    userLotOwner]

rowInsertLocation125CBake = 3
rowInsertLocation150CBake = 
if chamberChoice == '125C Bake':
    sheetBakeHTS.insert_rows(rowInsertLocation125CBake)


def insertData():
    for i in range(2, 10):
        cellref = sheetBakeHTS.cell(row=3, column=i)
        cellref.value = userDataList[i]
    wb.save('InsertRowTest.xlsx')
    print('Did this work?')


