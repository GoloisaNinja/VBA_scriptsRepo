Sub formatRoutebyAcct()
Dim ws As Worksheet
Dim lastRow As Long
Set ws = Sheets("ROUTED BY ACCT")
ws.Activate
ws.AutoFilter.ShowAllData
[B:B].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range("A1:M" & lastRow).Sort key1:=Range("B1:B" & lastRow), _
order1:=xlAscending, Header:=xlYes

Range("A1:M" & lastRow).Sort key1:=Range("A1:A" & lastRow), _
order1:=xlAscending, Header:=xlYes

[C:C].Select
With Selection
    .NumberFormat = "General"
    .Value = .Value
End With

Call emptyCustReport

End Sub
