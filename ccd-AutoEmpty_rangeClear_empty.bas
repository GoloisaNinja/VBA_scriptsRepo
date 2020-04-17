Sub CleanUp()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim ws As Worksheet
Dim DepWs As Worksheet
Dim upDateRte As Worksheet
Dim myRange As Range
Dim myRange2 As Range
Dim myDepRange As Range
Dim myUpDateRange1 As Range
Dim myUpDateRange2 As Range
Dim lastRow As Integer
Set ws = Sheets("ROUTED BY ACCT")
Set DepWs = Sheets("Routes With Departure")
Set upDateRte = Sheets("Updated Route Sheet")
'====================================================================
ws.Activate
ws.AutoFilter.ShowAllData

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Set myRange = ws.Range("A1:M" & lastRow)
Set myRange2 = ws.Range("AB2:AC" & lastRow)
myRange.Cells.ClearContents
myRange2.Cells.ClearContents
'ws.Range("A2").Value = "                                       >>>>>>>>>>>>>>>Make sure you Paste data in cell A2<<<<"
'===================================================================
DepWs.Activate
DepWs.AutoFilter.ShowAllData
Set myDepRange = DepWs.Range("A1:N" & lastRow)
myDepRange.Cells.ClearContents
'DepWs.Range("A2").Value = "                                   >>>>>>>>>>>>Make sure you Paste data in cell A2<<<<"
'====================================================================
upDateRte.Activate
upDateRte.AutoFilter.ShowAllData
Set myUpDateRange1 = upDateRte.Range("M2:R" & lastRow)
Set myUpDateRange2 = upDateRte.Range("Z2:AA" & lastRow)
myUpDateRange1.Cells.ClearContents
myUpDateRange2.Cells.ClearContents
'=====================================================================


ActiveWorkbook.Sheets("BUTTONS").Select

Sheets("BUTTONS").Activate
Sheets("BUTTONS").Range("D3").Value = Now

Application.ScreenUpdating = True

Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

