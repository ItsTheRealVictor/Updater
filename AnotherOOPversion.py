import openpyxl as opx
import pprint
import datetime
from dateutil import parser
from dateutil import relativedelta
import sys
import questionary as q
import pandas as pd

wb = opx.load_workbook(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx")
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

    def __init__(self, dataList):
        self.dataList = dataList

    dataList = []

    def getVariables(self):
        userLotNumber = q.text('Enter the Lot number: ')
        userPartNumber = q.text('Enter the part number: ')
        userNumOfLots = q.text('Enter the number of lots: ')
        userQuantity = q.text('Enter the quantity: ')
        userStartTime = q.text('Enter the starting time: ')
        userLotOwner = q.text('Enter the owner: ')
        userDateInput = q.text('Enter your start date (YYYY-MM-DD): ')
        userTimeIntervalInput = q.text('enter your time interval: ')

        varList = [userLotNumber, userPartNumber, userNumOfLots,
           userQuantity, userStartTime, userLotOwner,
           userDateInput, userTimeIntervalInput]

        for var in varList:
            item = var.ask()
            GetDataFromUser.dataList.append(item)

    def storeData(self):
        """Stores the data input by the user into a list (dataList)"""
        for index, variable in enumerate(GetDataFromUser.dataList):
            GetDataFromUser.dataList[index] = variable

    def computeDate(self, incomingDate):
        """Takes the date and time interval from the user and computes the future date and day. """
        userTimeIntervalDelta = relativedelta.relativedelta(
            hours=int(GetDataFromUser.dataList[-1])) # I don't know how to access the userTimeIntervalInput variable from the local
                                                     # scope of the function getVariables(). My workaround is to access it by slicing
                                                     # out of the list I created in the for loop at the end of getVariables()
        parsedUserDateInput = parser.parse(
            GetDataFromUser.dataList[6]).date() # Same as above, the item at dataList[6] is the user's input date. 
        future = parsedUserDateInput + userTimeIntervalDelta
        parsedFutureDate = datetime.datetime.strftime(future, "%Y-%m-%d")
        parsedUserDateInputDay = datetime.datetime.strftime(
            parsedUserDateInput, "%A")
        parsedFutureDay = datetime.datetime.strftime(future, "%A")
        GetDataFromUser.dataList.append(
            f"{parsedFutureDate} ({parsedFutureDay})")
        return f"{parsedFutureDate}" + f" {parsedFutureDay}"

    def dataSummary(self, ):
        """Gives the user a summary of the data they have enterd"""
        print(
            f"\nHere is a summary of the data you have entered \n{[i for i in GetDataFromUser.dataList]}")

    def doAllFunctions(self):
        GetDataFromUser.getVariables(self)
        GetDataFromUser.storeData(self)
        GetDataFromUser.computeDate(self, GetDataFromUser.dataList[6])
        # GetDataFromUser.dataSummary()


# GetDataFromUser.doAllFunctions()


class updateSpreadSheet:

    def __init__(self, chamberChoice, chamberList):
        self.chamberChoice = chamberChoice
        self.chamberList = chamberList

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


    def pickChamber(self):

        chamberDict = {k: v for k, v in enumerate(updateSpreadSheet.chamberList)} # I used enumerate in case I want to add or subtract chambers in the future.
        pprint.pprint(chamberDict)
        userChamberSelection = int(input("Enter the number which corresponds to the chamber: "))
        for k, v in chamberDict.items():
            if userChamberSelection == k:
                updateSpreadSheet.chamberChoice = f"{v}"

    inputDataList = [i for i in GetDataFromUser.dataList]
    inputDataList.insert(0, "")
    inputDataList.insert(1, "")
    insertLocation = 3
    for index, value in enumerate(chambersOnlySheetsList):
        if value == chamberChoice:
            wb[sheetsList[index]].insert_rows(insertLocation)
            for i in range(2, 10):
                cellref = wb[sheetsList[index]].cell(row=3, column=i)
                cellref.value = inputDataList[i]

class confirmUpdateSpreadsheet:

    def confirmUpdate():
        CHOICES = ['Update my spreadsheet and quit', 'Update my spreadsheet and add more parts', 'Quit']
        confirm = q.select("What would you like to do now?", choices = CHOICES).ask()
        if confirm == 'Update my spreadsheet and quit':
            wb.save(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx")
            sys.exit()
        elif confirm == 'Update my spreadsheet and add more parts':
            wb.save(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx")

        elif confirm == 'Quit':
            print("Bruh")
            wb.save(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx")

GetDataFromUser.doAllFunctions()
