Sub setPrevMonth()
Dim wb As Workbook
Dim mainS As Worksheet
Dim myRange1 As Range
Dim myRange2 As Range
Dim strPass As String

strPass = "holidaygolo04"
Set wb = ThisWorkbook
Set mainS = wb.Sheets("MAIN")
Set myRange1 = mainS.Range("C32:F37")
Set myRange2 = mainS.Range("C50:F55")

'unprotect so macro can run
mainS.Unprotect Password:=strPass
If Not Range("B6:B9").Locked = False Then
        mainS.Range("B6:B9").Locked = False
    End If
    
myRange2.Value2 = myRange1.Value2

'reprotect after value change
mainS.Protect Password:=strPass, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, UserInterfaceOnly:=True

End Sub
