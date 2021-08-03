
import openpyxl as opx
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from collections import defaultdict

wb = opx.Workbook()
refWB = opx.load_workbook(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx", data_only=True) #reference workbook, where data is pulled
refSheets = refWB.worksheets[0:14] #indices of the various chamber-specific sheets
activeSheet = wb.active
activeSheet.title = "Dashboard"


#color pallete templates, for changing cell colors
darkOrangeColor = PatternFill(patternType='solid', start_color='F6C010', end_color='FF5833')
lightOrangeColor = PatternFill(patternType='solid', start_color = 'FFE262', end_color='FFE262')
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
lotProgress = activeSheet['F2']
lotProgress.value = "Lot Progress"
lotProgressRefs = [chr(i) for i in range(ord('g'),ord('t')+1)]
for index, cell in enumerate(lotProgressRefs):
    activeSheet[f'{cell}3'] = refSheets[index].title
    activeSheet[f'{cell}3'].alignment = Alignment(wrap_text=True)


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

def getLotNums(sheetIndex):
    chamberLotNumList = list()
    for row in range(3, refSheets[sheetIndex].max_row):
        for column in "B":
            cellname = f'{column}{row}'
            cellValue = refSheets[sheetIndex][f'{cellname}'].value
            if cellValue is not None:
                chamberLotNumList.append((cellValue))
    return(chamberLotNumList)

#I've been stuck for a days on this so I'm resorting to a brute force way of solving this. I
#  want a list for each chamber, each list being the contents of row B from that chamber's sheet.
#It's ugly as hell but it works.
bake125ClotNames = getLotNums(0)
bake150CLotNames = getLotNums(1)
bake180CLotNames = getLotNums(2)
bake210CLotNames = getLotNums(3)
soak3060LotNames = getLotNums(4)
soak6060LotNames = getLotNums(5)
soak8585LotNames = getLotNums(6)
coldOven8LotNames = getLotNums(7)
uHast1LotNames = getLotNums(8)
uHast2LotNames = getLotNums(9)
uHast3LotNames = getLotNums(10)
uHast5LotNames = getLotNums(11)
TC2LotNames = getLotNums(12)
TC3LotNames = getLotNums(13)
# TC4LotNames = getLotNums(14)







chamberDict = defaultdict(list)


for i in refSheets:
    chamberDict[f'{i.title}'] #This initializes a dictionary with the chamber names as keys. The values will be the lists
                              #generated by the function getLotNums(). This function will be renamed later, it really truly
                              #acts as a way to get, from all sheets, the row values from a specific column. The dashboard
                              #will show the information from 
                                                                #column B(2), the lot number
                                                                #column M(13), the date/time of removal
                                                                #column N(14), the time until removal
                                                                #column O(15), the percentage of time remaining (0% means done)
print(chamberDict)
















# # for sheet in refSheets:
# for row in range(3, activeRefSheet.max_row):
#     for column in "B":
#         cellname = f'{column}{row}'
#         cellValue = activeRefSheet[f'{cellname}'].value
#         if cellValue is not None:
#             p.pprint((activeRefSheet.title, cellname))



wb.save(r"C:\Users\VD102541\Desktop\DashboardTest.xlsx")







