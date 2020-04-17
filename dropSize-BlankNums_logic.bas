Sub normalizeNums()
Dim wb As Workbook
Dim sh As Worksheet
Set wb = ThisWorkbook
Set sh = wb.Sheets("Sheet1")
Dim myRange As Range
Set myRange = sh.Range("b2:o6")
Dim cell

For Each cell In myRange
    If cell.Value <= 0 Then
    cell.Value = "" '
    End If
    Next cell
End Sub
