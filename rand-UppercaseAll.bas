Sub Upper_CaseAll()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Dim cell As Range
On Error Resume Next
For Each cell In Cells.SpecialCells(xlConstants, xlTextValues)
cell.Formula = UCase(cell.Formula)
Next
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

