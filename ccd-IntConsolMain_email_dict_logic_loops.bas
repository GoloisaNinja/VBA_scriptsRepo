Option Explicit


Public Function BuildHtmlBody_allroute()
    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww
    Dim v As Integer
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    v = 1
    
    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, delay, manualNotes

    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family: Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='font-family:Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    'html = html & "<table style='border-collapse:collapse'>"
    'html = html & "<table style = 'table-layout:fixed; width: 100%; white-space: nowrap; border : 1px solid black; cell padding =4'>"
    html = html & "<th style='font-color: white; padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Stop" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "City" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Cases" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Planned Departure Time" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Actual Departure Time" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Window" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Planned Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Est Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Actual Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Delay" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Notes" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 2 To lastRoww

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
        
        manualNotes = oSheet.Range("AW" & i).Value
        
        If IsError(manualNotes) Or manualNotes = 0 Then
        manualNotes = ""
        End If
        
        If custConc = lateTrip_allroute Then
        On Error Resume Next
            If v Mod 2 = 0 Then
                html = html & "<tr>"
                html = html & "<td style='font-size: 12px; padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & manualNotes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
                
            Else
            
                html = html & "<tr>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px; padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & manualNotes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
            End If
            
        v = v + 1
        'Range("AE" & i).Value = "Sent"
        'Range("AF" & i).Value = Range("S" & i).Value
        
        
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    BuildHtmlBody_allroute = html
End Function
Public Sub allRouteBlast()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim ws As Worksheet
Dim wsButtons As Worksheet
Dim wb As Workbook
Dim i As Long
Dim rw As Range
Dim fRw As Range
Dim conCount As Long
Dim outapp As Object
Dim outmail As Object
Dim lastRow As Long
Dim fLastRow As Long
Dim lateCount As Integer
Dim fLateCount As Integer
Set wb = ThisWorkbook
Set ws = wb.Sheets("ROUTED BY ACCT")
Set wsButtons = wb.Sheets("BUTTONS")
'ws.Activate
ws.AutoFilter.ShowAllData




lastRow = ws.Cells(Rows.Count, 33).End(xlUp).Row
fLastRow = ws.Cells(Rows.Count, 34).End(xlUp).Row
Set rw = ws.Range("AG2:AG" & lastRow)
Set fRw = ws.Range("AH2:AH" & fLastRow)
lateCount = Application.WorksheetFunction.Count(rw)
fLateCount = Application.WorksheetFunction.Count(fRw)


If intConAdHoc = False Then
    On Error Resume Next
    Application.OnTime alertTimeTwo, "allRouteBlast", , False
    alertTimeTwo = 0



If intConTimerTest <> "" Then
        alertTimeTwo = Now + TimeValue(intConTimerTest)
        Application.OnTime alertTimeTwo, "allRouteBlast", defIntConSchedUpper
        wsButtons.Range("C19").Value = intConTimerTest
        wsButtons.Range("C13").Value = alertTimeTwo
    End If
    If intConTimerTest = "" Then
        alertTimeTwo = Now + TimeValue(defaultIntCon)
'        Application.OnTime alertTimeThree, "dashBoardEmail", defDashSchedUpper
        alertTimeTwoAsNumber = alertTimeTwo - Int(alertTimeTwo)
        If alertTimeTwoAsNumber > defIntConSchedUpper Then
            Application.OnTime defIntConSched, "allRouteBlast", defIntConSchedUpper
            wsButtons.Range("C13").Value = defIntConSched
        Else
        Application.OnTime alertTimeTwo, "allRouteBlast", defIntConSchedUpper
        wsButtons.Range("C19").Value = defaultIntCon
        wsButtons.Range("C13").Value = alertTimeTwo
        End If
    End If




If lateCount = 0 And fLateCount = 0 Then
Exit Sub
End If
myArr_allroute = rw
myFranArr_allroute = fRw

myEndArr_allroute = RemoveDupesDict(myArr_allroute)



For i = LBound(myEndArr_allroute) To UBound(myEndArr_allroute)
    If myEndArr_allroute(i) <> "" And myEndArr_allroute(i) <> 0 Then
    
    lateTrip_allroute = myEndArr_allroute(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = buildRecips_allroute()
            If mainRecip_allroute = "" Then GoTo nextIterationOne
            .CC = ""
            .BCC = ""
            .Subject = "By Route 1st Tier Reporting - Concept " & lateTrip_allroute
            .HTMLBody = BuildHtmlBody_allroute()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIterationOne:
Next i



myFranEndArr_allroute = franRemoveDupesDict(myFranArr_allroute)



For i = LBound(myFranEndArr_allroute) To UBound(myFranEndArr_allroute)
    If myFranEndArr_allroute(i) <> "" And myFranEndArr_allroute(i) <> 0 Then
    
    myFranLateTrip_allroute = myFranEndArr_allroute(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = franBuildRecips_allroute()
            If mainRecip_allroute = "" Then GoTo nextIterationTwo
            .CC = ""
            .BCC = ""
            .Subject = "By Route 2nd Tier Reporting - Franchise " & myFranLateTrip_allroute
            .HTMLBody = franBuildHtmlBody_allroute()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIterationTwo:
Next i


'Sheets("BUTTONS").Activate
wsButtons.Range("P5").Value = Now


' If intConTimerTest <> "" Then
'        alertTimeTwo = Now + TimeValue(intConTimerTest)
'        Application.OnTime alertTimeTwo, "allRouteBlast", defIntConSchedUpper
'        wsButtons.Range("C19").Value = intConTimerTest
'        wsButtons.Range("C13").Value = alertTimeTwo
'    End If
'    If intConTimerTest = "" Then
'        alertTimeTwo = Now + TimeValue(defaultIntCon)
''        Application.OnTime alertTimeThree, "dashBoardEmail", defDashSchedUpper
'        alertTimeTwoAsNumber = alertTimeTwo - Int(alertTimeTwo)
'        If alertTimeTwoAsNumber > defIntConSchedUpper Then
'            Application.OnTime defIntConSched, "allRouteBlast", defIntConSchedUpper
'            wsButtons.Range("C13").Value = defIntConSched
'        Else
'        Application.OnTime alertTimeTwo, "allRouteBlast", defIntConSchedUpper
'        wsButtons.Range("C19").Value = defaultIntCon
'        wsButtons.Range("C13").Value = alertTimeTwo
'        End If
'    End If



'alertTimeTwo = Now + TimeValue("01:00:00")
'Application.OnTime alertTimeTwo, "allRouteBlast"


Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

Exit Sub

Else

If lateCount = 0 And fLateCount = 0 Then
Exit Sub
End If
myArr_allroute = rw
myFranArr_allroute = fRw

myEndArr_allroute = RemoveDupesDict(myArr_allroute)



For i = LBound(myEndArr_allroute) To UBound(myEndArr_allroute)
    If myEndArr_allroute(i) <> "" And myEndArr_allroute(i) <> 0 Then
    
    lateTrip_allroute = myEndArr_allroute(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = buildRecips_allroute()
            If mainRecip_allroute = "" Then GoTo nextIteration
            .CC = ""
            .BCC = ""
            .Subject = "By Route 1st Tier Reporting - Concept " & lateTrip_allroute
            .HTMLBody = BuildHtmlBody_allroute()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIteration:
Next i



myFranEndArr_allroute = franRemoveDupesDict(myFranArr_allroute)



For i = LBound(myFranEndArr_allroute) To UBound(myFranEndArr_allroute)
    If myFranEndArr_allroute(i) <> "" And myFranEndArr_allroute(i) <> 0 Then
    
    myFranLateTrip_allroute = myFranEndArr_allroute(i)
    
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = franBuildRecips_allroute()
            If mainRecip_allroute = "" Then GoTo nextIterations
            .CC = ""
            .BCC = ""
            .Subject = "By Route 2nd Tier Reporting - Franchise " & myFranLateTrip_allroute
            .HTMLBody = franBuildHtmlBody_allroute()
            .Send
        End With
    Set outmail = Nothing
    Set outapp = Nothing
End If
nextIterations:
Next i


'Sheets("BUTTONS").Activate
wsButtons.Range("P5").Value = Now

If intConTimerTest <> "" Then
        alertTimeTwo = Now + TimeValue(intConTimerTest)
        Application.OnTime alertTimeTwo, "allRouteBlast"
        wsButtons.Range("C19").Value = intConTimerTest
        wsButtons.Range("C13").Value = alertTimeTwo
    End If
    If intConTimerTest = "" Then
        alertTimeTwo = Now + TimeValue(defaultIntCon)
        Application.OnTime alertTimeTwo, "allRouteBlast"
        wsButtons.Range("C19").Value = defaultIntCon
        wsButtons.Range("C13").Value = alertTimeTwo
    End If
    

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End If
End Sub

Function RemoveDupesDict(myArr_allroute As Variant) As Variant

    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For i = LBound(myArr_allroute) To UBound(myArr_allroute)
            If IsMissing(myArr_allroute(i, 1)) = False Then
                .Item(myArr_allroute(i, 1)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function


Function buildRecips_allroute()
Dim oSheet As Worksheet
Dim wb As Workbook
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custConc
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
mainRecip_allroute = ""


For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custConc = oSheet.Range("AG" & i).Value
        primaryEmail = oSheet.Range("AK" & i).Value
        lString = InStr(mainRecip_allroute, primaryEmail)
            If lString = 0 And custConc = lateTrip_allroute And primaryEmail <> 0 And primaryEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & primaryEmail & ";"
            End If
        Next i

       
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custConc = oSheet.Range("AG" & i).Value
        secondaryEmail = oSheet.Range("AL" & i).Value
        mString = InStr(mainRecip_allroute, secondaryEmail)
            If mString = 0 And custConc = lateTrip_allroute And secondaryEmail <> 0 And secondaryEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & secondaryEmail & ";"
            End If
        Next i
        
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custConc = oSheet.Range("AG" & i).Value
        supEmail = oSheet.Range("AM" & i).Value
        nString = InStr(mainRecip_allroute, supEmail)
            If nString = 0 And custConc = lateTrip_allroute And supEmail <> 0 And supEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & supEmail & ";"
            End If
        Next i
buildRecips_allroute = mainRecip_allroute
End Function




Function franRemoveDupesDict(myFranArr_allroute As Variant) As Variant

    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    With d
        For i = LBound(myFranArr_allroute) To UBound(myFranArr_allroute)
            If IsMissing(myFranArr_allroute(i, 1)) = False Then
                .Item(myFranArr_allroute(i, 1)) = 1
            End If
        Next
        franRemoveDupesDict = .Keys
    End With
End Function

Function franBuildRecips_allroute()
Dim wb As Workbook
Dim oSheet As Worksheet
Set wb = ThisWorkbook
Set oSheet = wb.Sheets("ROUTED BY ACCT")
Dim i, lastRoww, primaryEmail, secondaryEmail, supEmail, lString, mString, nString, sendCheck, custFran
lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
mainRecip_allroute = ""

For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custFran = oSheet.Range("AH" & i).Value
        primaryEmail = oSheet.Range("AK" & i).Value
        lString = InStr(mainRecip_allroute, primaryEmail)
            If lString = 0 And custFran = myFranLateTrip_allroute And primaryEmail <> 0 And primaryEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & primaryEmail & ";"
            End If
        Next i

       
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custFran = oSheet.Range("AH" & i).Value
        secondaryEmail = oSheet.Range("AL" & i).Value
        mString = InStr(mainRecip_allroute, secondaryEmail)
            If mString = 0 And custFran = myFranLateTrip_allroute And secondaryEmail <> 0 And secondaryEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & secondaryEmail & ";"
            End If
        Next i
        
For i = 2 To lastRoww

        sendCheck = oSheet.Range("V" & i).Value
        custFran = oSheet.Range("AH" & i).Value
        supEmail = oSheet.Range("AM" & i).Value
        nString = InStr(mainRecip_allroute, supEmail)
            If nString = 0 And custFran = myFranLateTrip_allroute And supEmail <> 0 And supEmail <> "" Then
                mainRecip_allroute = mainRecip_allroute & supEmail & ";"
            End If
        Next i

franBuildRecips_allroute = mainRecip_allroute
End Function

Public Function franBuildHtmlBody_allroute()
    Dim wb As Workbook
    Dim oSheet As Worksheet
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Dim i, lastRoww
    Dim v As Integer
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    v = 1

    Dim html, custName, custRoute, custCity, custCases, planArrival, estArrival, custConc, custFran, sendEmail, actArrival, driver, deptTime, tripStop, planDepTime, window, delay, manualNotes

    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<table style='font-family:Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    'html = html & "<table style='border-collapse:collapse'>"
    'html = html & "<table style = 'table-layout:fixed; width: 100%; white-space: nowrap; border : 1px solid black; cell padding =4'>"
    html = html & "<tr>"
    html = html & "<th style='font-color: white; padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Route" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Stop" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Customer" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "City" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Cases" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Planned Departure Time" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Actual Departure Time" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Window" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Planned Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Est Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Actual Arrival" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Delay" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #c5cedb;border-color: black; border-width: 1px 1px 0 0;'>" & "Notes" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 2 To lastRoww

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
        
        manualNotes = oSheet.Range("AW" & i).Value
        
        If IsError(manualNotes) Or manualNotes = 0 Then
        manualNotes = ""
        End If
        
        
        If custFran = myFranLateTrip_allroute Then
        On Error Resume Next
        If v Mod 2 = 0 Then
                html = html & "<tr>"
                html = html & "<td style='font-size: 12px; padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                html = html & "<td style='font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & manualNotes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
                
            Else
            
                html = html & "<tr>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px; padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custRoute & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & tripStop & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custName & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCity & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & custCases & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planDepTime & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & deptTime & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & window & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & planArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & estArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & actArrival & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & delay & "</td>"
                html = html & "<td style='background-color: #f2f3f4; font-size: 12px;padding: 5px; border-style: solid; border-color: black; border-width: 1px 1px 0 0;'>" & manualNotes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
            End If
        v = v + 1
        
        
        
        
    End If
    Next i
        

    html = html & "</table></div></body></html>"
    franBuildHtmlBody_allroute = html
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

