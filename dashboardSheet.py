
import openpyxl as opx
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from collections import defaultdict
import pprint as p
import datetime as dt
import pandas as pd

wb = opx.Workbook()
# reference workbook, where data is pulled
refWB = opx.load_workbook(
    r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx", data_only=True)
# indices of the various chamber-specific sheets
refSheets = refWB.worksheets[1:15]
activeSheet = wb.active
activeSheet.title = "Dashboard"


#color pallete templates, for changing cell colors
darkRedColor = PatternFill(patternType='solid', start_color='DE3163', end_color='DE3163')
darkOrangeColor = PatternFill(patternType='solid', start_color='F6C010', end_color='FF5833')
lightOrangeColor = PatternFill(patternType='solid', start_color='FFE262', end_color='FFE262')
greenColor = PatternFill(patternType='solid', start_color='43B854', end_color='43B854')
lightGreenColor = PatternFill(patternType='solid', start_color='76FF91', end_color='76FF91')
blueColor = PatternFill(patternType='solid', start_color='5A63FF', end_color='5A63FF')
lightBlueColor = PatternFill(patternType='solid', start_color='AEC1FF', end_color='AEC1FF')


#Cell with Today's Date
activeSheet['A1'] = "Today's date is "
activeSheet['A1'].alignment = Alignment(wrap_text=True)
activeSheet['B1'] = "=NOW()"
activeSheet['B1'].number_format = "MM/DD/YYYY"

#column widths
activeSheet.column_dimensions['B'].width = 15
activeSheet.column_dimensions['C'].width = 25
activeSheet.column_dimensions['D'].width = 20
activeSheet.column_dimensions['F'].width = 25


#Title cell
currentProgress = activeSheet['C3']
currentProgress.value = "Lots currently in progress: "
currentProgress.fill = darkOrangeColor
for cell in activeSheet.iter_rows(min_row=4, max_row=17, min_col=3, max_col=3):
    cell[0].fill = lightOrangeColor

#Total number of units in chambers
totalUnits = activeSheet['D3']
totalUnits.value = "Total units in progress"
totalUnits.fill = blueColor
for cell in activeSheet.iter_rows(min_row=4, max_row=17, min_col=4, max_col=4):
    cell[0].fill = lightBlueColor

#Lot Progress
lotProgress = activeSheet['F3']
lotProgress.value = "Lot Progress"
lotProgress.fill = darkRedColor
lotProgress.font = Font(color="FDFEFE")
lotProgressRefs = [f"F{i}" for i in range(4, 18)]
for index, cell in enumerate(lotProgressRefs):
    activeSheet[f'{cell}'] = refSheets[index].title
    activeSheet[f'{cell}'].alignment = Alignment(wrap_text=True)


#Chamber Cells
chamberCellRefs = [(f"B{i}") for i in range(4, 18)]
for index, cell in enumerate(chamberCellRefs):
    activeSheet[f'{cell}'] = refSheets[index].title
    activeSheet[f'{cell}'].fill = greenColor
    activeSheet[f'{cell}'].alignment = Alignment(wrap_text=True)

#First Progress Chart - num of lots, total num of units for each chamber
lots125CBake = refWB['Dashboard']['C3'].value
activeSheet['C4'].value = lots125CBake


#EXPERIMENTAL CODE

def getData(sheetIndex, columnIndex):

    """Takes the integer index of the sheet (from the list refSheets) as the first argument, the column index (as a letter) for the
       second argument"""

    dataList = list()
    for row in range(3, refSheets[sheetIndex].max_row):
        for column in columnIndex:
            cellname = f'{column}{row}'
            cellValue = refSheets[sheetIndex][f'{cellname}'].value
            if cellValue is not None:
                dataList.append((cellValue))
    return(dataList)

chamberDict = defaultdict(list) #initiates the dictionary which will hold all of our dynamic dashboard data. The keys
                                #are chambers, the values will be lists containing data from columns B(lot number),
                                #M(removal date/time), N(time until removal) and O(percentage time remaining)

for sheet in refSheets:    #Establishes chambers (i.e. titles of the worksheets) as dictionary keys
    chamberDict[f'{sheet.title}']

for index, sheet in enumerate(refSheets): #Gathers the lists of and puts it into the dictionary
    chamberDict[f'{sheet.title}'].append(getData(index, "B")) #Lot Nums (i.e. RA numbers)
    chamberDict[f'{sheet.title}'].append(getData(index, "L")) #Date/time in
    chamberDict[f'{sheet.title}'].append(getData(index, "M")) #Removal date
    chamberDict[f'{sheet.title}'].append(getData(index, "N")) #Time until removal
    chamberDict[f'{sheet.title}'].append(getData(index, "O")) # % of time remaining


for sheet in refSheets:
    (chamberDict[f'{sheet.title}'][-1]) = [ f'{round(x * 100, 1)}%' for x in (chamberDict[f'{sheet.title}'][-1])]

zippedData = zip(chamberDict['125C Bake'])

#I'm thinking to make a pandas dataframe with the data from the dictionary. Columns will be part number, date/time in, 
# date/time out, remaining time, and % time remaining. Each chamber will have it's own dataframe? That might work. 

columnList = ["RA Number", "Date/Time in", "Removal Date/Time", "Time Until Removal", r"% of time remaining"]
df = pd.DataFrame([chamberDict['125C Bake'][i] for i in range(0, 5)], index=columnList)

# this kind of does what I want, the output is transposed though. I want 5 columns of 8 rows instead of 5 rows of 8 columns. 
# I still need to  -
                        # figure out how to transpose this to the dimensions I want
                        # Write a function which makes a dataframe out of each key/lists pair in the chamberDict dictionary
                        # Figure out how to insert those dataframes into the corresponding spots in the dashboardSheet






wb.save(r"C:\Users\VD102541\Desktop\DashboardTest.xlsx")


