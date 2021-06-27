import openpyxl as opx
import pprint
import datetime
from dateutil import parser
from dateutil import relativedelta
import sys

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
        userTimeIntervalDelta = relativedelta.relativedelta(hours=int(GetDataFromUser.userTimeIntervalInput))
        parsedUserDateInput = parser.parse(GetDataFromUser.userDateInput).date()
        future = parsedUserDateInput + userTimeIntervalDelta
        parsedFutureDate = datetime.datetime.strftime(future, "%Y-%m-%d")
        parsedUserDateInputDay = datetime.datetime.strftime(parsedUserDateInput, "%A")
        parsedFutureDay = datetime.datetime.strftime(future, "%A")
        GetDataFromUser.dataList.append(f"{parsedFutureDate} ({parsedFutureDay})")

    def dataSummary():
        """Gives the user a summary of the data they have enterd"""
        print(f"\nHere is a summary of the data you have entered \n{[i for i in GetDataFromUser.dataList]}")


    def doAllFunctions():
        GetDataFromUser.storeData()
        GetDataFromUser.computeDate(GetDataFromUser.userDateInput)
        GetDataFromUser.dataSummary()

GetDataFromUser.doAllFunctions()

class UpdateSpreadSheet:
    pass

