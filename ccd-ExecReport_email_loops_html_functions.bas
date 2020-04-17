Option Explicit
Public Function BuildHtmlBody_summaryDash()
    Dim oSheet As Worksheet
    Dim buttonSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("Route Summary")
    Set buttonSheet = wb.Sheets("BUTTONS")
    Dim i, lastRoww, realLastRow
    Dim v As Integer
    oSheet.AutoFilter.ShowAllData
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    realLastRow = 2
    v = 1
    Dim html, trip, driver, schedDisp, actDep, depStatus, status, late, onStop, total, notes
    Dim totalTrips As Integer
    Dim totalMissed As Integer
    Dim totalManual As Integer
    Dim totalSitting As Integer
    Dim combManAndSitting As Integer
    Dim totalEarly As Integer
    Dim totalNotTracking As Integer
    Dim totalLate As Integer
    Dim schDispAsString As String
    Dim actDepAsString As String
    Dim totalComplete As String
    
    Dim percentMissed As Integer
    Dim percentManual As Integer
    Dim percentEarly As Integer
    Dim percentNoData As Integer
    Dim percentLate As Integer
    
    
    For i = 3 To lastRoww
        If oSheet.Range("A" & i).Value <> 0 Then
        realLastRow = realLastRow + 1
        End If
    Next i
    
    
    
    totalTrips = Application.WorksheetFunction.CountA(oSheet.Range("A3:A" & realLastRow))
    totalMissed = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "Missed")
    totalManual = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "Manual")
    totalSitting = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "Sitting")
    totalEarly = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "Early")
    totalNotTracking = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "No Data")
    totalLate = Application.WorksheetFunction.CountIf(oSheet.Range("H3:H" & realLastRow), "Late")
    totalComplete = Application.WorksheetFunction.CountIf(oSheet.Range("I3:I" & realLastRow), "Complete")
    
    combManAndSitting = (totalManual + totalSitting)
    
    percentMissed = (totalMissed / totalTrips) * 100
    percentEarly = (totalEarly / totalTrips) * 100
    percentNoData = (totalNotTracking / totalTrips) * 100
    percentLate = (totalLate / totalTrips) * 100
    percentManual = (combManAndSitting / totalTrips) * 100
    
    

    html = "<!DOCTYPE html><body><html>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<h1 style='text-align: left; font-family: Arial;'>" & "Routing Summary Dashboard" & "</h1>"
    html = html & "<h2 style='text-align: left; font-family: Arial;'>" & WeekdayName(Weekday(Now)) & " " & Now & "</h2>"
    html = html & "<h2 style='text-align: left; font-family: Arial;'>" & "Live Telogis Pull Timestamp: " & buttonSheet.Range("A3").Value & "</h2>"
    html = html & "<h3 style ='text-align: left; font-family: Arial;'>" & "Total Trips: " & totalTrips & " " & " " & " || " & " " & " " & "Completed Trips: " & totalComplete & "</h3>"
    html = html & "<h3 style ='text-align: left; font-family: Arial;'>" & "Total Early: " & "(" & totalEarly & ")" & " " & percentEarly & "%" & " " & "  " & " || " & " " & " " & "Total Late: " & "(" & totalLate & ")" & " " & percentLate & "%" & " " & "  " & " || " & " " & " " & "Total Manual: " & "(" & combManAndSitting & ")" & " " & percentManual & "%" & " " & "</h3>"
    'html = html & "<h3>" & "Total Late: " & totalLate & " " & "(" & percentLate & "%" & ")" & "</h3>"
    html = html & "<h3 style ='text-align: left; font-family: Arial;'>" & "Total Missed: " & "(" & totalMissed & ")" & " " & percentMissed & "%" & " " & "  " & " || " & " " & " " & "Total Not Tracking: " & "(" & totalNotTracking & ")" & " " & percentNoData & "%" & " " & "</h3>"
    'html = html & "<h3>" & "Total Not Tracking: " & totalNotTracking & " " & "(" & percentNoData & "%" & ")" & "</h3>"
    html = html & "<table align = 'left'>"
    html = html & "<table style='font-family: Arial; border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; border-width: 0 0 1px 1px;'>"
    html = html & "<tr>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Trip" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Driver" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Scheduled Dispatch" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Actual Departure" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Departure Status" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Last Stop Status" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Condition of Trip" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "On Stop" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Total Stops" & "</th>"
    html = html & "<th style='font-color: white;padding: 5px; border-style: solid; background-color: #0033FF;border-color: #ccc; border-width: 1px 1px 0 0;'>" & "Notes" & "</th>"
    html = html & "</tr>"
    ' Build a html table based on rows data
    For i = 3 To realLastRow
    
        

        trip = oSheet.Range("A" & i).Value
        
        If IsError(trip) Then
        trip = ""
        End If
        
        driver = oSheet.Range("B" & i).Value
        
        If IsError(driver) Then
        driver = ""
        End If
        
        schedDisp = oSheet.Range("E" & i)
        
        If IsError(schedDisp) Then
        schedDisp = ""
        End If
        
        actDep = oSheet.Range("F" & i)
        
        If IsError(actDep) Then
        actDep = ""
        End If
        
        depStatus = oSheet.Range("G" & i)
        
        If IsError(depStatus) Then
        depStatus = ""
        End If
        
        status = oSheet.Range("H" & i).Value
        
        If IsError(status) Then
        status = ""
        End If
        
        late = oSheet.Range("I" & i).Value
        
        If IsError(late) Then
        late = ""
        End If
        If IsNumeric(late) Then
        late = Format(late, "hh:mm:ss")
        End If
        
        onStop = oSheet.Range("J" & i).Value
        
        If IsError(onStop) Then
        onStop = ""
        End If
        
        total = oSheet.Range("K" & i).Value
        
        If IsError(total) Then
        total = ""
        End If
        
        notes = oSheet.Range("L" & i).Value
        
        If IsError(notes) Or notes = 0 Then
        notes = ""
        End If
        
        schDispAsString = Format(schedDisp, "hh:mm AM/PM")
        actDepAsString = Format(actDep, "hh:mm AM/PM")
    
        
        On Error Resume Next
            If v Mod 2 = 0 Then
                html = html & "<tr>"
                html = html & "<td style='font-size: 12px; padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & trip & "</td>"
                html = html & "<td style='font-size: 12px; padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & schDispAsString & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actDepAsString & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & depStatus & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & status & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & late & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & onStop & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & total & "</td>"
                html = html & "<td style='font-size: 12px; text-align: center;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & notes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
        
            Else
        
                html = html & "<tr>"
                html = html & "<td style='font-size: 12px; background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & trip & "</td>"
                html = html & "<td style='font-size: 12px; background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & driver & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & schDispAsString & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & actDepAsString & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & depStatus & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & status & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & late & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & onStop & "</td>"
                html = html & "<td style='font-size: 12px; text-align: right;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & total & "</td>"
                html = html & "<td style='font-size: 12px; text-align: center;background-color: #cac2c0;padding: 5px; border-style: solid; border-color: #ccc; border-width: 1px 1px 0 0;'>" & notes & "</td>"
                On Error Resume Next
                html = html & "</tr>"
            
            End If
            v = v + 1
            
'
        
    
    Next i
        

    html = html & "</table></div></body></html>"
    BuildHtmlBody_summaryDash = html
End Function


Public Sub executiveDash()
Dim wb As Workbook
Dim ws As Worksheet
Dim outapp As Object
Dim outmail As Object
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
On Error Resume Next


If dashAdHoc = False Then
    On Error Resume Next
    Application.OnTime alertTimeThree, "executiveDash", , False
    alertTimeThree = 0



    If customTimerTest <> "" Then
        alertTimeThree = Now + TimeValue(customTimerTest)
        Application.OnTime alertTimeThree, "executiveDash", defDashSchedUpper
        ws.Range("C21").Value = customTimerTest
        ws.Range("C15").Value = alertTimeThree
    End If
    If customTimerTest = "" Then
        alertTimeThree = Now + TimeValue(defaultDash)
    '   Application.OnTime alertTimeThree, "dashBoardEmail", defDashSchedUpper
        alertTimeThreeAsNumber = alertTimeThree - Int(alertTimeThree)
                If alertTimeThreeAsNumber > defDashSchedUpper Then
                    Application.OnTime defDashSched, "executiveDash", defDashSchedUpper
                    ws.Range("C15").Value = defDashSched
                Else
                    Application.OnTime alertTimeThree, "executiveDash", defDashSchedUpper
                    ws.Range("C21").Value = defaultDash
                    ws.Range("C15").Value = alertTimeThree
                End If
    End If




Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = "patrick.doyle@maines.net"
            .CC = ""
            .BCC = "jarrett.newby@maines.net;marc.craig@maines.net;jonathan.collins@maines.net"
            .Subject = "Summary Routing Dashboard " & Now
            '.Attachments.Add "C:\Users\Jonathan Collins\Pictures\quickViewDash.png"
            .HTMLBody = BuildHtmlBody_summaryDash()
            .Send
        End With

    
    Set outmail = Nothing
    Set outapp = Nothing

ws.Range("P14").Value = Now

Exit Sub

Else

Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = "patrick.doyle@maines.net"
            .CC = ""
            .BCC = "jarrett.newby@maines.net;marc.craig@maines.net;jonathan.collins@maines.net"
            .Subject = "Summary Routing Dashboard " & Now
            '.Attachments.Add "C:\Users\Jonathan Collins\Pictures\quickViewDash.png"
            .HTMLBody = BuildHtmlBody_summaryDash()
            .Send
        End With

    
    Set outmail = Nothing
    Set outapp = Nothing

ws.Range("P14").Value = Now

If customTimerTest <> "" Then
        alertTimeThree = Now + TimeValue(customTimerTest)
        Application.OnTime alertTimeThree, "executiveDash"
        ws.Range("C21").Value = customTimerTest
        ws.Range("C15").Value = alertTimeThree
    End If
    If customTimerTest = "" Then
        alertTimeThree = Now + TimeValue(defaultDash)
        Application.OnTime alertTimeThree, "executiveDash"
        ws.Range("C21").Value = defaultDash
        ws.Range("C15").Value = alertTimeThree
    End If

End If

End Sub


