Option Explicit

Public Sub WorkBook_BeforeClose(Cancel As Boolean)
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
ws.Range("C18").Value = defaultTel
ws.Range("C19").Value = defaultIntCon
ws.Range("C21").Value = defaultDash

Call disableTimerEvents




End Sub

Public Sub Workbook_Open()
'Dim alertTime
'Dim alertTimeTwo
'Dim alertTimeThree
Dim answer As String
Dim wb As Workbook
Dim ws As Worksheet
Dim wsRoute As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
Set wsRoute = wb.Sheets("ROUTED BY ACCT")

Call allSheetProtect

answer = MsgBox("Do you want to enable Scheduled Events?", vbYesNo, "Engage Schedule")
'answer = InputBox(Prompt:="Yes/No", Title:="Timer Enable?", Default:="Please enter Yes or No")
'If answer = "No" Then
If answer = vbNo Then
Call declareAlerts

ws.Range("C12").Value = alertTimeOne
ws.Range("C13").Value = alertTimeTwo
ws.Range("C14").Value = alertTimeFour
ws.Range("C15").Value = alertTimeThree

ws.Range("G12") = defTelSched
ws.Range("H12") = defTelSchedUpper
ws.Range("I12") = restoreTelSchedUpper
ws.Range("C18") = defaultTel
ws.Range("G18") = telAdHoc

ws.Range("G13") = defIntConSched
ws.Range("H13") = defIntConSchedUpper
ws.Range("I13") = restoreIntConSchedUpper
ws.Range("C19") = defaultIntCon
ws.Range("G19") = intConAdHoc

ws.Range("G15") = defDashSched
ws.Range("H15") = defDashSchedUpper
ws.Range("I15") = restoreDashSchedUpper
ws.Range("C21") = defaultDash
ws.Range("G21") = dashAdHoc

Exit Sub
End If

Call declareAlerts

'alertTimeOne = Now + TimeValue("00:05:00")
'Application.OnTime alertTimeOne, "Main"
'
'alertTimeTwo = Now + TimeValue("00:010:00")
'Application.OnTime alertTimeTwo, "allRouteBlast"
'
'alertTimeThree = Now + TimeValue("00:06:00")
'Application.OnTime alertTimeThree, "dashBoardEmail"
'
'alertTimeFour = Now + TimeValue("00:05:00")
'Application.OnTime alertTimeFour, "newDynamicRoute"

Application.OnTime defTelSched, "Main", defTelSchedUpper
ws.Range("C12") = defTelSched
ws.Range("G12") = defTelSched
ws.Range("H12") = defTelSchedUpper
ws.Range("C18") = defaultTel
ws.Range("G18") = telAdHoc
ws.Range("I12") = restoreTelSchedUpper

Application.OnTime defIntConSched, "allRouteBlast", defIntConSchedUpper
ws.Range("C13") = defIntConSched
ws.Range("G13") = defIntConSched
ws.Range("H13") = defIntConSchedUpper
ws.Range("C19") = defaultIntCon
ws.Range("G19") = intConAdHoc
ws.Range("I13") = restoreIntConSchedUpper

Application.OnTime defDashSched, "executiveDash", defDashSchedUpper
ws.Range("C15") = defDashSched
ws.Range("G15") = defDashSched
ws.Range("H15") = defDashSchedUpper
ws.Range("C21") = defaultDash
ws.Range("G21") = dashAdHoc
ws.Range("I15") = restoreDashSchedUpper

'ws.Range("C18").Value = alertTimeOne
'ws.Range("C19").Value = alertTimeTwo
'ws.Range("C20").Value = alertTimeFour
'ws.Range("C21").Value = alertTimeThree

End Sub




