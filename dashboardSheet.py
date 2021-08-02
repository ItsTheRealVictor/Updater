
import openpyxl as opx
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

wb = opx.Workbook()
refWB = opx.load_workbook(r"C:\Users\valex\Desktop\VictorChamberUsageWithProgBar.xlsx")
refSheets = refWB.worksheets[1:15]
activeSheet = wb.active
activeSheet.title = "Dashboard"

#Cell with Today's Date
activeSheet['A1'] = "Today's date is "
activeSheet['A1'].alignment = Alignment(wrap_text=True)
activeSheet['B1'] = "=NOW()"
activeSheet['B1'].number_format = "MM/DD/YYYY"
activeSheet.column_dimensions['B'].width = 15
activeSheet.column_dimensions['C'].width = 25

orangeColor = PatternFill(patternType='solid',
                       start_color='F6C010', end_color='F6C010')
greenColor = PatternFill(patternType='solid',
                         start_color='43B854', end_color='43B854')


#Title cell
currentProgress = activeSheet['C3']
currentProgress.value = "Lots currently in progress: "
currentProgress.fill = orangeColor

#Chamber Cells
chamberCellRefs = [(f"B{i}") for i in range(4, 17)]
for index, cell in enumerate(chamberCellRefs):
    activeSheet[f'{cell}'] = refSheets[index].title
    activeSheet[f'{cell}'].fill = greenColor
    activeSheet[f'{cell}'].alignment = Alignment(wrap_text=True)




wb.save(r"C:\Users\valex\Desktop\DashboardTest.xlsx")
