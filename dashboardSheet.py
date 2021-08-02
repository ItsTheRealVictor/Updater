
import openpyxl as opx
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

wb = opx.Workbook()
refWB = opx.load_workbook(r"C:\Users\VD102541\Desktop\VictorChamberUsageWithProgBar.xlsx") #reference workbook, where data is pulled
refSheets = refWB.worksheets[1:15] #indices of the various chamber-specific sheets
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
for cell in activeSheet.iter_rows(min_row=4, max_row=16, min_col=3, max_col=3):
    cell[0].fill = lightOrangeColor

#Total number of units in chambers
totalUnits = activeSheet['D3']
totalUnits.value = "Total units in progress"
totalUnits.fill = blueColor
for cell in activeSheet.iter_rows(min_row=4, max_row=16, min_col=4, max_col=4):
    cell[0].fill = lightBlueColor

#Lot Progress
lotProgress = activeSheet['F2']
lotProgress.value = "Lot Progress"
lotProgressRefs = [chr(i) for i in range(ord('g'),ord('s')+1)]
for index, cell in enumerate(lotProgressRefs):
    activeSheet[f'{cell}3'] = refSheets[index].title
    activeSheet[f'{cell}3'].alignment = Alignment(wrap_text=True)



#Chamber Cells
chamberCellRefs = [(f"B{i}") for i in range(4, 17)]
for index, cell in enumerate(chamberCellRefs):
    activeSheet[f'{cell}'] = refSheets[index].title
    activeSheet[f'{cell}'].fill = greenColor
    activeSheet[f'{cell}'].alignment = Alignment(wrap_text=True)




wb.save(r"C:\Users\VD102541\Desktop\DashboardTest.xlsx")







