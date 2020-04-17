Option Explicit
Public Sub disableTimerEvents()
Dim wb
Dim ws
Dim myMsgBox
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
On Error Resume Next


'If alertTimeThree = 0 And alertTimeOne = 0 And alertTimeTwo = 0 And alertTimeFour = 0 Then
'Exit Sub
'End If

myMsgBox = MsgBox("Warning, you are disabling all timers!", vbOKCancel, "TIMERS WILL BE DISABLED")
If myMsgBox = vbCancel Then
Exit Sub
End If
MsgBox "You disabled all timers successfully.", vbOKOnly, "Successful Cancelled Timers"

    
        Application.OnTime alertTimeOne, "Main", , False
        Application.OnTime defTelSched, "Main", , False
        Application.OnTime alertTimeTwo, "allRouteBlast", , False
        Application.OnTime defIntConSched, "allRouteBlast", , False
        Application.OnTime alertTimeThree, "executiveDash", , False
        Application.OnTime defDashSched, "executiveDash", , False
        Application.OnTime alertTimeFour, "newDynamicRoute", , False
    

finalStage:
alertTimeOne = 0
alertTimeTwo = 0
alertTimeThree = 0
alertTimeFour = 0

telAdHoc = False
intConAdHoc = False
extLateAdHoc = False
dashAdHoc = False

telTimerTest = ""
intConTimerTest = ""
extLateTimerTest = ""
customTimerTest = ""


ws.Range("A6").Value = Now

ws.Range("C12").Value = alertTimeOne
ws.Range("C13").Value = alertTimeTwo
ws.Range("C14").Value = alertTimeFour
ws.Range("C15").Value = alertTimeThree

ws.Range("C18").Value = defaultTel
ws.Range("C19").Value = defaultIntCon
ws.Range("C20").Value = defaultExtLate
ws.Range("C21").Value = defaultDash

ws.Range("g18").Value = telAdHoc
ws.Range("g19").Value = intConAdHoc
ws.Range("g20").Value = extLateAdHoc
ws.Range("g21").Value = dashAdHoc

End Sub


