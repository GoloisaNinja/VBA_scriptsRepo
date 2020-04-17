Option Explicit

Public Function BuildHtmlBody_newdynamic()
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
    'html = html & "<th style='padding: 10px; border-style: solid; background-color: #cac2c0;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Window" & "</th>"
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
        
        window = oSheet.Range("I" & i).Value
        
        If IsError(window) Then
        window = ""
        End If
        
        
        If custConc = lateTrip_newdynamic And sendEmail = "YES" Then
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
        
'        If ((secEmailCheck = "" Or secEmailCheck = 0) And (supEmailCheck = "" Or supEmailCheck = 0)) Then
'        oSheet.Range("AB" & i).Value = "Sent"
'        oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
'        End If
        
        oSheet.Range("AB" & i).Value = "Sent"
        oSheet.Range("AC" & i).Value = oSheet.Range("R" & i).Value
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    BuildHtmlBody_newdynamic = html
End Function
Public Sub newDynamicRoute()

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
Set buttonsWS = wb.Sheets("BUTTONS")
ws.AutoFilter.ShowAllData
lastRow = ws.Cells(Rows.Count, 33).End(xlUp).Row
fLastRow = ws.Cells(Rows.Count, 34).End(xlUp).Row
custLastRow = ws.Cells(Rows.Count, 3).End(xlUp).Row
Set rw = ws.Range("AG2:AG" & lastRow)
Set fRw = ws.Range("AH2:AH" & fLastRow)
Set custHSrw = ws.Range("C2:C" & custLastRow)
lateCount = Application.WorksheetFunction.Count(rw)
fLateCount = Application.WorksheetFunction.Count(fRw)



If lateCount = 0 And fLateCount = 0 Then
Exit Sub
End If
myArr_newdynamic = rw
myFranArr_newdynamic = fRw
myCustArr_newdynamic = custHSrw

myEndArr_newdynamic = RemoveDupesDict(myArr_newdynamic)



For i = LBound(myEndArr_newdynamic) To UBound(myEndArr_newdynamic)
    If myEndArr_newdynamic(i) <> "" And myEndArr_newdynamic(i) <> 0 Then
    
    lateTrip_newdynamic = myEndArr_newdynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = buildRecips_newdynamic()
            If mainRecip_newdynamic = "" Then GoTo nextIteration
            .CC = ""
            .BCC = ""
            .Subject = "1st Tier Reporting - Delay Concept " & lateTrip_newdynamic
            .HTMLBody = BuildHtmlBody_newdynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIteration:
Next i



myFranEndArr_newdynamic = franRemoveDupesDict(myFranArr_newdynamic)



For i = LBound(myFranEndArr_newdynamic) To UBound(myFranEndArr_newdynamic)
    If myFranEndArr_newdynamic(i) <> "" And myFranEndArr_newdynamic(i) <> 0 Then
    
    myFranLateTrip_newdynamic = myFranEndArr_newdynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = franBuildRecips_newdynamic()
            If mainRecip_newdynamic = "" Then GoTo nextIterations
            .CC = ""
            .BCC = ""
            .Subject = "2nd Tier Reporting - Delay Franchise " & myFranLateTrip_newdynamic
            .HTMLBody = franBuildHtmlBody_newdynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIterations:
Next i


myCustEndArr_newdynamic = custRemoveDupesDict(myCustArr_newdynamic)



For i = LBound(myCustEndArr_newdynamic) To UBound(myCustEndArr_newdynamic)
    If myCustEndArr_newdynamic(i) <> "" And myCustEndArr_newdynamic(i) <> 0 Then
    
    myCustLateTrip_newdynamic = myCustEndArr_newdynamic(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = custBuildRecips_newdynamic()
            If mainRecip_newdynamic = "" Then GoTo nextCustIterations
            .CC = ""
            .BCC = ""
            .Subject = "3rd Tier Reporting - Delay Customer " & myCustLateTrip_newdynamic
            .HTMLBody = custBuildHtmlBody_newdynamic()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextCustIterations:
Next i




buttonsWS.Range("R8").Value = Now
'
'alertTimeFour = Now + TimeValue("00:05:00")
'Application.OnTime alertTimeFour, "newDynamicRoute"
'
'buttonsWS.Range("C14").Value = alertTimeFour
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


Function RemoveDupesDict(myArr_newdynamic As Variant) As Variant

    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For i = LBound(myArr_newdynamic) To UBound(myArr_newdynamic)
            If IsMissing(myArr_newdynamic(i, 1)) = False Then
                .Item(myArr_newdynamic(i, 1)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function


Function buildRecips_newdynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custConc
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        mainRecip_newdynamic = ""
        whenLateCheck = 0

For i = 2 To lastRoww


        sendCheck = oSheet.Range("V" & i).Value
        custConc = oSheet.Range("AG" & i).Value
        primaryEmail = oSheet.Range("X" & i).Value
        secondaryEmail = oSheet.Range("Y" & i).Value
        supEmail = oSheet.Range("Z" & i).Value
        lString = InStr(mainRecip_newdynamic, primaryEmail)
            If lString = 0 And custConc = lateTrip_newdynamic And primaryEmail <> 0 And primaryEmail <> "" And sendCheck = "YES" Then
                mainRecip_newdynamic = mainRecip_newdynamic & primaryEmail & ";"
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
        
buildRecips_newdynamic = mainRecip_newdynamic
End Function




Function franRemoveDupesDict(myFranArr_newdynamic As Variant) As Variant

    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For i = LBound(myFranArr_newdynamic) To UBound(myFranArr_newdynamic)
            If IsMissing(myFranArr_newdynamic(i, 1)) = False Then
                .Item(myFranArr_newdynamic(i, 1)) = 1
            End If
        Next
        franRemoveDupesDict = .Keys
    End With
End Function

Function custRemoveDupesDict(myCustArr_newdynamic As Variant) As Variant

    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For i = LBound(myCustArr_newdynamic) To UBound(myCustArr_newdynamic)
            If IsMissing(myCustArr_newdynamic(i, 1)) = False Then
                .Item(myCustArr_newdynamic(i, 1)) = 1
            End If
        Next
        custRemoveDupesDict = .Keys
    End With
End Function

Function franBuildRecips_newdynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custFran
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row

mainRecip_newdynamic = ""

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
        mString = InStr(mainRecip_newdynamic, secondaryEmail)
            If mString = 0 And custFran = myFranLateTrip_newdynamic And sendCheck = "YES" And secondaryEmail <> 0 And secondaryEmail <> "" Then
                mainRecip_newdynamic = mainRecip_newdynamic & secondaryEmail & ";"
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
        
franBuildRecips_newdynamic = mainRecip_newdynamic
End Function

Function custBuildRecips_newdynamic()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custCust
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
mainRecip_newdynamic = ""

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
        nString = InStr(mainRecip_newdynamic, supEmail)
            If nString = 0 And custCust = myCustLateTrip_newdynamic And sendCheck = "YES" And supEmail <> 0 And supEmail <> "" Then
                mainRecip_newdynamic = mainRecip_newdynamic & supEmail & ";"
            End If
        Next i
        
custBuildRecips_newdynamic = mainRecip_newdynamic
End Function

Public Function franBuildHtmlBody_newdynamic()
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
        
        
        If custFran = myFranLateTrip_newdynamic And sendEmail = "YES" Then
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
    franBuildHtmlBody_newdynamic = html
End Function

Public Function custBuildHtmlBody_newdynamic()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row

    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, custCust, delay

    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:Helvetica; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
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
        
        
        If custCust = myCustLateTrip_newdynamic And sendEmail = "YES" Then
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
    custBuildHtmlBody_newdynamic = html
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

