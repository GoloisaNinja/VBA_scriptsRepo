Sub ExtractTabsXLS()

Dim mySH As Worksheet

For Each mySH In ActiveWorkbook.Worksheets
If mySH.Visible = xlSheetVisible Then
mySH.Copy
With ActiveWorkbook

'//try to insert protection here
ActiveSheet.Protect Password:="0621", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
AllowFormattingColumns:=True, AllowFormattingRows:=True
'//end protect section

.SaveAs Filename:="C:\Users\Jonathan Collins\Documents\Locations\" & mySH.Name & ".xlsx"
.Close
End With
End If
Next mySH

End Sub
