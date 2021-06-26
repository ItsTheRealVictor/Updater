import openpyxl as opx
import pprint
import datetime
from dateutil import parser
from dateutil import relativedelta

class GetDataFromUser:

    def __init__(self,
                    userLotNumber,
                    userPartNumber,
                    userNumOfLots,
                    userQuantity,
                    userStartTime,
                    userLotOwner):
        self.userLotNumber = userLotNumber
        self.userPartNumber = userPartNumber
        self.userNumOfLots = userNumOfLots
        self.userQuantity = userQuantity
        self.userStartTime = userStartTime
        self.userLotOwner = userLotOwner

#     userLotNumber = input('Enter the Lot number: ')


# userPartNumber = input('Enter the part number: ')
# userNumOfLots = input('Enter the number of lots: ')
# userQuantity = input('Enter the quantity: ')
# userStartTime = input('Enter the starting time: ')
# userLotOwner = input('Enter the owner: ')
