Option Explicit
Public myEmptyCusArr() As Variant

Sub emptyCustReport()
Dim wb As Workbook
Dim ws As Worksheet
Dim lastRow As Long
Dim i, v
Dim outapp As Object
Dim outmail As Object
Set wb = ThisWorkbook
Set ws = wb.Sheets("ROUTED BY ACCT")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Application.DisplayAlerts = False
Application.ScreenUpdating = False
ws.AutoFilter.ShowAllData

ReDim Preserve myEmptyCusArr(0)
For v = 2 To lastRow
        If ((ws.Range("AK" & v).Value = "") Or (ws.Range("AK" & v).Value = 0)) And ((ws.Range("AL" & v).Value = "") Or (ws.Range("AL" & v).Value = 0)) And ((ws.Range("AM" & v).Value = "") Or (ws.Range("AM" & v).Value = 0)) Then
            If UBound(myEmptyCusArr) >= 0 Then
            myEmptyCusArr(UBound(myEmptyCusArr)) = ws.Range("C" & v).Value
            ReDim Preserve myEmptyCusArr(0 To UBound(myEmptyCusArr) + 1)
            End If
        End If
    Next v
If IsNull(myEmptyCusArr) Or IsEmpty(myEmptyCusArr) Or IsArray(myEmptyCusArr) = False Then
Exit Sub
End If
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = "jarrett.newby@maines.net;jonathan.collins@maines.net"
            .CC = ""
            .BCC = ""
            .Subject = "CCD Customer Maintenance Run " & Now
            .HTMLBody = buildCusMainEmail()
            .Send
        End With
    
    Set outmail = Nothing
    Set outapp = Nothing
Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub


Public Function buildCusMainEmail()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww, n
    Dim v As Integer
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    v = 1
    Dim html, custName, custRoute, custNum
    html = "<!DOCTYPE html><body><html>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<h4 style='text-align: left; font-family: Arial;'>" & "The below customers require maintenance - primary/secondary/supp email fields are empty." & "</h4>"
    html = html & "<table style='font-family: Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    html = html & "<tr>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Customer Number" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"

    html = html & "</tr>"
    ' Build a html table based on rows data
    
For n = LBound(myEmptyCusArr) To UBound(myEmptyCusArr)

    
    For i = 2 To lastRoww

        custRoute = oSheet.Range("A" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        custName = oSheet.Range("D" & i).Value
        
        If IsError(custName) Then
        custName = ""
        End If
        
        custNum = oSheet.Range("C" & i).Value
        
        If IsError(custNum) Then
        custNum = ""
        End If
        
        
        
        
        If custNum = myEmptyCusArr(n) Then
        On Error Resume Next
            If v Mod 2 = 0 Then
                html = html & "<tr>"
                html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custNum & "</td>"
                html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                On Error Resume Next
                html = html & "</tr>"
        
            Else
        
                html = html & "<tr>"
                html = html & "<td style='font-size: 11px; background-color: #f2f3f4;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='font-size: 11px; background-color: #f2f3f4;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custNum & "</td>"
                html = html & "<td style='font-size: 11px; background-color: #f2f3f4;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                On Error Resume Next
                html = html & "</tr>"
            
            End If
            v = v + 1
            

        
    End If
    Next i
Next n

    html = html & "</table></div></body></html>"
    buildCusMainEmail = html
End Function
