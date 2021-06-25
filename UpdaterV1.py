import datetime
from dateutil import parser
from dateutil import relativedelta
import openpyxl as opx
import pprint

wb = opx.load_workbook(r"C:\Users\VD102541\Desktop\Copy of West Lab Chamber Usage WW25 2021.xlsx")



userLotNumber = input('Enter the Lot number: ')
userPartNumber = input('Enter the part number: ')
userNumOfLots = input('Enter the number of lots: ')
userQuantity = input('Enter the quantity: ')
userStartTime = input('Enter the starting time: ')
userLotOwner = input('Enter the owner: ')
userDataList = [userLotNumber, userPartNumber, userNumOfLots, userQuantity, userStartTime, userLotOwner]

#Get the location of the lab (ADD THIS FUNCTION LATER)

# labLocation = ''
# userChoice = input("Choose your lab: \n(1): East Lab\n(2): West Lab\n")
# if userChoice == "1":
#  labLocation = "West Lab"
# elif userChoice == "2":
#  labLocation = "East Lab"


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


userDateInput = input('Enter your start date: ')
parsedUserDateInput = parser.parse(userDateInput).date()
parsedUserDateInputDay = datetime.datetime.strftime(parsedUserDateInput, "%A")

userTimeIntervalInput = input('enter your time interval: ')
userTimeIntervalDelta = relativedelta.relativedelta(hours=int(userTimeIntervalInput))


def calculateFutureDate(incoming):
    future = parsedUserDateInput + userTimeIntervalDelta
    parsedFutureDate = datetime.datetime.strftime(future, "%Y-%m-%d")
    parsedFutureDay = datetime.datetime.strftime(future, "%A")
    return f"{parsedFutureDate} ({parsedFutureDay})"

sheetBakeHTS = wb['Bake_HTS']

# Test functions to find the best way of putting the data from the user into the spreadsheet

userDataList = ['', '', 
                    userLotNumber, 
                    userPartNumber, 
                    userNumOfLots, 
                    userQuantity, 
                    parsedUserDateInput, 
                    calculateFutureDate(userDateInput), 
                    userStartTime, 
                    userLotOwner]


def insertData():
    sheetBakeHTS.insert_rows(3)
    for i in range(2, 10):
        cellref = sheetBakeHTS.cell(row=3, column=i)
        cellref.value = userDataList[i]
    wb.save('InsertRowTest.xlsx')
    print('Did this work?')





if __name__ == '__main__':
    def confirmAllInfo():
        print(f"""You have entered:
        Lot Number = {userLotNumber}
        Part Number = {userPartNumber}
        Quantity = {userQuantity}
        Start Date = {userDateInput} ({parsedUserDateInputDay})
        Time Interval = {userTimeIntervalInput}
        End Date = {calculateFutureDate(userDateInput)}
        Start Time = {userStartTime}
        Owner = {userLotOwner}""")
        
    print(confirmAllInfo())
    insert125CBakeRow()
    wb.save('InsertRowTest.xlsx')
    print('Did this work?')