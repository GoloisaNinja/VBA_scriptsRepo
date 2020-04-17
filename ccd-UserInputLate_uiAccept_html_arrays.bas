Option Explicit
Public myArray_uidynamic As Variant
Public myEndArray_uidynamic As Variant
Public Function BuildHtmlBody_uidynamic()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww, primEmailCheck, secEmailCheck, supEmailCheck
    Dim v As Integer
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    v = 1
    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, delay

    html = "<!DOCTYPE html><body><html>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='font-family:Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    'html = html & "<style = ""text/css""> tr:nth-child(even){background-color: #cac2c0;}</style>"'
    'html = html & "<table style='border-collapse:collapse'>"
    'html = html & "<table style = 'table-layout:fixed; width: 100%; white-space: nowrap; border : 1px solid black; cell padding =4'>"
    html = html & "<tr>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Stop" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "City" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Cases" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Departure Time" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Departure Time" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Window" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Est Arrival" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Delay" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 2 To lastRoww
    
        primEmailCheck = oSheet.Range("X" & i).Value
        secEmailCheck = oSheet.Range("Y" & i).Value
        supEmailCheck = oSheet.Range("Z" & i).Value

        custRoute = oSheet.Range("A" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        custName = oSheet.Range("D" & i).Value
        
        If IsError(custName) Then
        custName = ""
        End If
        
        custCity = oSheet.Range("F" & i).Value
        
        If IsError(custCity) Then
        custCity = ""
        End If
        
        custCases = oSheet.Range("J" & i).Value
        
        If IsError(custCases) Then
        custCases = ""
        End If
        
        planArrival = oSheet.Range("K" & i).Value
        
        If IsError(planArrival) Then
        planArrival = ""
        End If
        
        estArrival = oSheet.Range("Q" & i).Value
        
        If IsError(estArrival) Then
        estArrival = ""
        End If
        
        actArrival = oSheet.Range("O" & i).Value
        
        If IsError(actArrival) Then
        actArrival = ""
        End If
        
        custConc = oSheet.Range("AG" & i).Value
        
        If IsError(custConc) Then
        custConc = ""
        End If
        
        custFran = oSheet.Range("AH" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        sendEmail = oSheet.Range("V" & i).Value
        
        If IsError(sendEmail) Then
        sendEmail = ""
        End If
        
        delay = oSheet.Range("AJ" & i).Value
        
        If IsError(delay) Then
        delay = ""
        End If
        
        driver = oSheet.Range("L" & i).Value
        
        If IsError(driver) Then
        driver = ""
        End If
        
        deptTime = oSheet.Range("AO" & i).Value
        
        If IsError(deptTime) Then
        deptTime = ""
        End If
        
        tripStop = oSheet.Range("AP" & i).Value
        
        If IsError(tripStop) Then
        tripStop = ""
        End If
        
        planDepTime = oSheet.Range("AR" & i).Value
        
        If IsError(planDepTime) Then
        planDepTime = ""
        End If
        
        window = oSheet.Range("I" & i).Value
        
        If IsError(window) Then
        window = ""
        End If
        
        
        If custConc = lateTrip_uidynamic And sendEmail = "YES" Then
        On Error Resume Next
            If v Mod 2 = 0 Then
                html = html & "<tr>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                On Error Resume Next
                html = html & "</tr>"
        
            Else
        
                html = html & "<tr>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                On Error Resume Next
                html = html & "</tr>"
            
            End If
            v = v + 1
            
'            If ((secEmailCheck = "" Or secEmailCheck = 0) And (supEmailCheck = "" Or supEmailCheck = 0)) Then
'                oSheet.Range("AB" & i).Value = "Sent"
'                oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
'            End If
            
            oSheet.Range("AB" & i).Value = "Sent"
            oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    BuildHtmlBody_uidynamic = html
End Function
Public Sub newUserInputDynamicRoute()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim ws As Worksheet
Dim wb As Workbook
Dim buttonsWS As Worksheet
Dim i As Long
Dim rw As Range
Dim fRw As Range
Dim custHSrw As Range
Dim custLastRow As Long
Dim conCount As Long
Dim outapp As Object
Dim outmail As Object
Dim lastRow As Long
Dim fLastRow As Long
Dim lateCount As Integer
Dim fLateCount As Integer
Set wb = ThisWorkbook
Set ws = wb.Sheets("ROUTED BY ACCT")
ws.AutoFilter.ShowAllData

Set buttonsWS = wb.Sheets("BUTTONS")


lastRow = ws.Cells(Rows.Count, 33).End(xlUp).Row
fLastRow = ws.Cells(Rows.Count, 34).End(xlUp).Row
custLastRow = ws.Cells(Rows.Count, 3).End(xlUp).Row
Set rw = ws.Range("AG2:AG" & lastRow)
Set fRw = ws.Range("AH2:AH" & fLastRow)
Set custHSrw = ws.Range("C2:C" & custLastRow)
lateCount = Application.WorksheetFunction.Count(rw)
fLateCount = Application.WorksheetFunction.Count(fRw)



'If lateCount = 0 And fLateCount = 0 Then
'Exit Sub
'End If
'myArr = rw
'myFranArr = fRw
'myCustArr = custHSrw
'
'myEndArr = RemoveDupesDict(myArr)


myArray_uidynamic = Application.InputBox("List Concept in the following format: {con1, con2, con3, ...}", Type:=64)
If IsArray(myArray_uidynamic) <> False Then




For i = LBound(myArray_uidynamic) To UBound(myArray_uidynamic)
    If myArray_uidynamic(i) <> "" And myArray_uidynamic(i) <> 0 Then
    
    lateTrip_uidynamic = myArray_uidynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = buildRecips_uidynamic()
            If mainRecip_uidynamic = "" Then GoTo nextIteration
            .CC = ""
            .BCC = ""
            .Subject = "1st Tier Reporting - Delay Concept " & lateTrip_uidynamic
            .HTMLBody = BuildHtmlBody_uidynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIteration:
Next i







For i = LBound(myArray_uidynamic) To UBound(myArray_uidynamic)
    If myArray_uidynamic(i) <> "" And myArray_uidynamic(i) <> 0 Then
    
    myFranLateTrip_uidynamic = myArray_uidynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = franBuildRecips_uidynamic()
            If mainRecip_uidynamic = "" Then GoTo nextIterations
            .CC = ""
            .BCC = ""
            .Subject = "2nd Tier Reporting - Delay Franchise " & myFranLateTrip_uidynamic
            .HTMLBody = franBuildHtmlBody_uidynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIterations:
Next i





For i = LBound(myArray_uidynamic) To UBound(myArray_uidynamic)
    If myArray_uidynamic(i) <> "" And myArray_uidynamic(i) <> 0 Then
    
    myCustLateTrip_uidynamic = myArray_uidynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = custBuildRecips_uidynamic()
            If mainRecip_uidynamic = "" Then GoTo nextCustIterations
            .CC = ""
            .BCC = ""
            .Subject = "3rd Tier Reporting - Delay Customer " & myCustLateTrip_uidynamic
            .HTMLBody = custBuildHtmlBody_uidynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextCustIterations:
Next i




buttonsWS.Range("R11").Value = Now

Else
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Exit Sub
End If
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub
Function buildRecips_uidynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custConc
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        mainRecip_uidynamic = ""

For i = 2 To lastRoww


        sendCheck = oSheet.Range("V" & i).Value
        custConc = oSheet.Range("AG" & i).Value
        primaryEmail = oSheet.Range("X" & i).Value
        lString = InStr(mainRecip_uidynamic, primaryEmail)
            If lString = 0 And custConc = lateTrip_uidynamic And primaryEmail <> 0 And primaryEmail <> "" And sendCheck = "YES" Then
                mainRecip_uidynamic = mainRecip_uidynamic & primaryEmail & ";"
            End If
        Next i

       
'For i = 2 To lastRoww
'
'
'        sendCheck = Range("V" & i).Value
'        custConc = Range("AG" & i).Value
'        secondaryEmail = Range("AL" & i).Value
'        mString = InStr(mainRecip, secondaryEmail)
'            If mString = 0 And custConc = lateTrip And secondaryEmail <> "" And secondaryEmail <> 0 And sendCheck = "YES" Then
'                mainRecip = mainRecip & secondaryEmail & ";"
'            End If
'        Next i
'
'For i = 2 To lastRoww
'
'
'        sendCheck = Range("V" & i).Value
'        custConc = Range("AG" & i).Value
'        supEmail = Range("AM" & i).Value
'        nString = InStr(mainRecip, supEmail)
'            If nString = 0 And custConc = lateTrip And supEmail <> "" And supEmail <> 0 And sendCheck = "YES" Then
'                mainRecip = mainRecip & supEmail & ";"
'            End If
'        Next i
        
buildRecips_uidynamic = mainRecip_uidynamic
End Function
Function franBuildRecips_uidynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custFran
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row

mainRecip_uidynamic = ""

'For i = 2 To lastRoww
'
'        sendCheck = Range("V" & i).Value
'        custFran = Range("AH" & i).Value
'        'primaryEmail = Range("X" & i).Value
'        primaryEmail = Range("AK" & i).Value
'        lString = InStr(mainRecip, primaryEmail)
'            If lString = 0 And custFran = myFranLateTrip And sendCheck = "YES" And primaryEmail <> 0 And primaryEmail <> "" Then
'                mainRecip = mainRecip & primaryEmail & ";"
'            End If
'        Next i

       
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custFran = oSheet.Range("AH" & i).Value
        'secondaryEmail = Range("Y" & i).Value
        secondaryEmail = oSheet.Range("Y" & i).Value
        mString = InStr(mainRecip_uidynamic, secondaryEmail)
            If mString = 0 And custFran = myFranLateTrip_uidynamic And sendCheck = "YES" And secondaryEmail <> 0 And secondaryEmail <> "" Then
                mainRecip_uidynamic = mainRecip_uidynamic & secondaryEmail & ";"
            End If
        Next i
        
'For i = 2 To lastRoww
'
'        sendCheck = Range("V" & i).Value
'        custFran = Range("AH" & i).Value
'        'supEmail = Range("Z" & i).Value
'        supEmail = Range("AM" & i).Value
'        nString = InStr(mainRecip, supEmail)
'            If nString = 0 And custFran = myFranLateTrip And sendCheck = "YES" And supEmail <> 0 And supEmail <> "" Then
'                mainRecip = mainRecip & supEmail & ";"
'            End If
'        Next i
        
franBuildRecips_uidynamic = mainRecip_uidynamic
End Function
Function custBuildRecips_uidynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custCust
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
mainRecip_uidynamic = ""

'For i = 2 To lastRoww
'
'        sendCheck = Range("V" & i).Value
'        custCust = Range("AH" & i).Value
'        'primaryEmail = Range("X" & i).Value
'        primaryEmail = Range("AK" & i).Value
'        lString = InStr(mainRecip, primaryEmail)
'            If lString = 0 And custFran = myFranLateTrip And sendCheck = "YES" And primaryEmail <> 0 And primaryEmail <> "" Then
'                mainRecip = mainRecip & primaryEmail & ";"
'            End If
'        Next i
'
'
'For i = 2 To lastRoww
'
'        sendCheck = Range("V" & i).Value
'        custCust = Range("AH" & i).Value
'        'secondaryEmail = Range("Y" & i).Value
'        secondaryEmail = Range("AL" & i).Value
'        mString = InStr(mainRecip, secondaryEmail)
'            If mString = 0 And custFran = myFranLateTrip And sendCheck = "YES" And secondaryEmail <> 0 And secondaryEmail <> "" Then
'                mainRecip = mainRecip & secondaryEmail & ";"
'            End If
'        Next i
        
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custCust = oSheet.Range("C" & i).Value
        'supEmail = Range("Z" & i).Value
        supEmail = oSheet.Range("Z" & i).Value
        nString = InStr(mainRecip_uidynamic, supEmail)
            If nString = 0 And custCust = myCustLateTrip_uidynamic And sendCheck = "YES" And supEmail <> 0 And supEmail <> "" Then
                mainRecip_uidynamic = mainRecip_uidynamic & supEmail & ";"
            End If
        Next i
        
custBuildRecips_uidynamic = mainRecip_uidynamic
End Function
Public Function franBuildHtmlBody_uidynamic()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww, primEmailCheck, secEmailCheck, supEmailCheck
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row

    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, delay

    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='font-family:Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    'html = html & "<table style='border-collapse:collapse'>"
    'html = html & "<table style = 'table-layout:fixed; width: 100%; white-space: nowrap; border : 1px solid black; cell padding =4'>"
    html = html & "<tr>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Stop" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "City" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Cases" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Departure Time" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Departure Time" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Arrival" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Est Arrival" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Arrival" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Delay" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 2 To lastRoww
    
    
        primEmailCheck = oSheet.Range("X" & i).Value
        secEmailCheck = oSheet.Range("Y" & i).Value
        supEmailCheck = oSheet.Range("Z" & i).Value
            

        custRoute = oSheet.Range("A" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        custName = oSheet.Range("D" & i).Value
        
        If IsError(custName) Then
        custName = ""
        End If
        
        custCity = oSheet.Range("F" & i).Value
        
        If IsError(custCity) Then
        custCity = ""
        End If
        
        custCases = oSheet.Range("J" & i).Value
        
        If IsError(custCases) Then
        custCases = ""
        End If
        
        planArrival = oSheet.Range("K" & i).Value
        
        If IsError(planArrival) Then
        planArrival = ""
        End If
        
        estArrival = oSheet.Range("Q" & i).Value
        
        If IsError(estArrival) Then
        estArrival = ""
        End If
        
        actArrival = oSheet.Range("O" & i).Value
        
        If IsError(actArrival) Then
        actArrival = ""
        End If
        
        custConc = oSheet.Range("AG" & i).Value
        
        If IsError(custConc) Then
        custConc = ""
        End If
        
        custFran = oSheet.Range("AH" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        sendEmail = oSheet.Range("V" & i).Value
        
        If IsError(sendEmail) Then
        sendEmail = ""
        End If
        
        delay = oSheet.Range("AJ" & i).Value
        
        If IsError(delay) Then
        delay = ""
        End If
        
        driver = oSheet.Range("L" & i).Value
        
        If IsError(driver) Then
        driver = ""
        End If
        
        deptTime = oSheet.Range("AO" & i).Value
        
        If IsError(deptTime) Then
        deptTime = ""
        End If
        
        tripStop = oSheet.Range("AP" & i).Value
        
        If IsError(tripStop) Then
        tripStop = ""
        End If
        
        planDepTime = oSheet.Range("AR" & i).Value
        
        If IsError(planDepTime) Then
        planDepTime = ""
        End If
        
        
        If custFran = myFranLateTrip_uidynamic And sendEmail = "YES" Then
        On Error Resume Next
        html = html & "<tr>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & delay & "</td>"
        On Error Resume Next
        html = html & "</tr>"
        
        
'        If supEmailCheck = 0 Or supEmailCheck = "" Then
'        oSheet.Range("AB" & i).Value = "Sent"
'        oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
'        End If
        
        oSheet.Range("AB" & i).Value = "Sent"
        oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    franBuildHtmlBody_uidynamic = html
End Function
Public Function custBuildHtmlBody_uidynamic()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row

    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, custCust, delay

    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='font-family:Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    'html = html & "<table style='border-collapse:collapse'>"
    'html = html & "<table style = 'table-layout:fixed; width: 100%; white-space: nowrap; border : 1px solid black; cell padding =4'>"
    html = html & "<tr>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Stop" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "City" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Cases" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Departure Time" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Departure Time" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Window" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Planned Arrival" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Est Arrival" & "</th>"
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Arrival" & "</th>"
    html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Delay" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 2 To lastRoww

        custRoute = oSheet.Range("A" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        custCust = oSheet.Range("C" & i).Value
        
        If IsError(custCust) Then
        custCust = ""
        End If
        
        custName = oSheet.Range("D" & i).Value
        
        If IsError(custName) Then
        custName = ""
        End If
        
        custCity = oSheet.Range("F" & i).Value
        
        If IsError(custCity) Then
        custCity = ""
        End If
        
        custCases = oSheet.Range("J" & i).Value
        
        If IsError(custCases) Then
        custCases = ""
        End If
        
        planArrival = oSheet.Range("K" & i).Value
        
        If IsError(planArrival) Then
        planArrival = ""
        End If
        
        estArrival = oSheet.Range("Q" & i).Value
        
        If IsError(estArrival) Then
        estArrival = ""
        End If
        
        actArrival = oSheet.Range("O" & i).Value
        
        If IsError(actArrival) Then
        actArrival = ""
        End If
        
        custConc = oSheet.Range("AG" & i).Value
        
        If IsError(custConc) Then
        custConc = ""
        End If
        
        custFran = oSheet.Range("AH" & i).Value
        
        If IsError(custRoute) Then
        custRoute = ""
        End If
        
        sendEmail = oSheet.Range("V" & i).Value
        
        If IsError(sendEmail) Then
        sendEmail = ""
        End If
        
        delay = oSheet.Range("AJ" & i).Value
        
        If IsError(delay) Then
        delay = ""
        End If
        
        driver = oSheet.Range("L" & i).Value
        
        If IsError(driver) Then
        driver = ""
        End If
        
        deptTime = oSheet.Range("AO" & i).Value
        
        If IsError(deptTime) Then
        deptTime = ""
        End If
        
        tripStop = oSheet.Range("AP" & i).Value
        
        If IsError(tripStop) Then
        tripStop = ""
        End If
        
        planDepTime = oSheet.Range("AR" & i).Value
        
        If IsError(planDepTime) Then
        planDepTime = ""
        End If
        
        window = oSheet.Range("I" & i).Value
        
        If IsError(window) Then
        window = ""
        End If
        
        
        If custCust = myCustLateTrip_uidynamic And sendEmail = "YES" Then
        On Error Resume Next
        html = html & "<tr>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custName & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & window & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
        'html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
        html = html & "<td style='padding: 10px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & delay & "</td>"
        On Error Resume Next
        html = html & "</tr>"
        
        oSheet.Range("AB" & i).Value = "Sent"
        oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
        
        
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    custBuildHtmlBody_uidynamic = html
End Function

'concept, layout, inception, integration, and selling everyone we could do it by Marc Craig (Dungeon Master +50 to Fire Damage)
'transportation consults, impossible understanding of the ridiculous data,logical sheet operation, primary,secondary, and thirdary framework and telogis wizardry by Jarrett Newby (Ninja Wizard +50 to Critical Hit)
'telogis API and all associated vb code supplied by Phil Deckers (Glorious Paladin +20 HP)
'vb code and custom functions by Jon Collins (tavern drunkard #4 w/ missing tunic -400 to party morale +60 chance of goblin attacks)
'speed enhancements by Sonic the Hedgehog
' __          __
'   \__(..)__/
'
'MainesSolutions
'admin@mainessolutions.net
'


