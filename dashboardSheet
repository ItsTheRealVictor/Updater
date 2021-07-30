import openpyxl as opx
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

wb = opx.Workbook()
refWB = opx.load_workbook(r"C:\Users\VD102541\Desktop\VictorChamberUsage.xlsx")
refSheets = refWB.worksheets[0:14]

dashBoard = wb.create_sheet("Dashboard")
wb.active = wb['Dashboard'] 
activeSheet = wb.active


Bake = refWB['125C Bake']
currentProgress = activeSheet['B3']
currentProgress.value = "Lots currently in progress: "
redColor = PatternFill(patternType='solid', start_color='F6C010', end_color='F6C010')
for cell in activeSheet.iter_cols(min_row=3, max_row=3, min_col=2, max_col=4):
    cell[0].fill = redColor



wb.save(r"C:\Users\VD102541\Desktop\DashboardTest.xlsx")





