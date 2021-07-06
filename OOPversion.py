import openpyxl as opx
import pprint
import datetime
from dateutil import parser
from dateutil import relativedelta
import sys
import pandas as pd

wb = opx.load_workbook(r"C:\Users\valex\Desktop\Updater\VictorFartChamberUsage.xlsx")
# # List of sheets
sheetsList=wb.sheetnames
chambersOnlySheetsList=sheetsList[0:14]
writer = pd.ExcelWriter('TestFart.xlsx', engine='openpyxl')
writer.book = wb



class OpeningScreen:

    while True:
        openingScreen = input("""
        ##############################
        UPDATER
        Version 1.0
        by Victor Delgado
        ---------
        Press 'c' to continue
        Press 'q' to quit
        ##############################
        """)
        if openingScreen == 'c':
            break
        elif openingScreen == 'q':
            sys.exit()
        else:
            print("Invalid input. Try Again")


class GetDataFromUser:

    """Starting with a group of variables that will hold data taken from the user about the
        parts, chamber, and their computed future date."""
    userLotNumber = input('Enter the Lot number: ')
    userPartNumber = input('Enter the part number: ')
    userNumOfLots = input('Enter the number of lots: ')
    userQuantity = input('Enter the quantity: ')
    userStartTime = input('Enter the starting time: ')
    userLotOwner = input('Enter the owner: ')
    userDateInput = input('Enter your start date (YYYY-MM-DD): ')
    userTimeIntervalInput = input('enter your time interval: ')

    dataList = [userLotNumber, userPartNumber, userNumOfLots, userQuantity, userStartTime,
                userLotOwner, userDateInput, userTimeIntervalInput]

    def storeData():
        """Stores the data input by the user into a list (dataList)"""
        for index, variable in enumerate(GetDataFromUser.dataList):
            GetDataFromUser.dataList[index] = variable

    def computeDate(incomingDate):
        """Takes the date and time interval from the user and computes the future date and day. """
        userTimeIntervalDelta = relativedelta.relativedelta(
            hours=int(GetDataFromUser.userTimeIntervalInput))
        parsedUserDateInput = parser.parse(
            GetDataFromUser.userDateInput).date()
        future = parsedUserDateInput + userTimeIntervalDelta
        parsedFutureDate = datetime.datetime.strftime(future, "%Y-%m-%d")
        parsedUserDateInputDay = datetime.datetime.strftime(
            parsedUserDateInput, "%A")
        parsedFutureDay = datetime.datetime.strftime(future, "%A")
        GetDataFromUser.dataList.append(
            f"{parsedFutureDate} ({parsedFutureDay})")
        return f"{parsedFutureDate}" + f" {parsedFutureDay}"

    def dataSummary():
        """Gives the user a summary of the data they have enterd"""
        print(
            f"\nHere is a summary of the data you have entered \n{[i for i in GetDataFromUser.dataList]}")

    def doAllFunctions():
        GetDataFromUser.storeData()
        GetDataFromUser.computeDate(GetDataFromUser.userDateInput)
        # GetDataFromUser.dataSummary()


GetDataFromUser.doAllFunctions()

#chamberChoice is an empty variable because the choice has yet to made by the user
chamberChoice = ''
chamberList = [
        "125C Bake",
        "150C Bake",
        "180C Bake",
        "210C Bake",
        "30.60 Soak",
        "60.60 Soak",
        "85.85 Soak""Oven 8 Cold",
        "UHAST 1",
        "UHAST 2",
        "UHAST 5",
        "TC.2_D",
        "TC.3_D",
        "TC.4_D"
                        ]

chamberDict = {k: v for k, v in enumerate(chamberList)} # I used enumerate in case I want to add or subtract chambers in the future.
pprint.pprint(chamberDict)
userChamberSelection = int(
            input("Enter the number which corresponds to the chamber: "))


for k, v in chamberDict.items():
    if userChamberSelection == k:
        chamberChoice = f"{v}"

inputDataList = ['','', GetDataFromUser.userLotNumber,
                    GetDataFromUser.userPartNumber,
                    GetDataFromUser.userNumOfLots,
                    GetDataFromUser.userQuantity,
                    GetDataFromUser.userDateInput,
                    GetDataFromUser.computeDate(GetDataFromUser.userDateInput),
                    GetDataFromUser.userStartTime,
                    GetDataFromUser.userLotOwner]
insertLocation = 3
for index, value in enumerate(chambersOnlySheetsList):
    if value == chamberChoice:
        wb[sheetsList[index]].insert_rows(insertLocation)
        for i in range(2, 10):
            cellref = wb[sheetsList[index]].cell(row=3, column=i)
            cellref.value = inputDataList[i]
wb.save(r"C:\Users\valex\Desktop\Updater\VictorFartChamberUsage.xlsx")
