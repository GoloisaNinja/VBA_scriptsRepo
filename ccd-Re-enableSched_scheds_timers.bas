Option Explicit

Public Sub reEnableTimerEvents()
Dim wb
Dim ws
Dim myMsgBox
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
If alertTimeThree <> 0 And alertTimeOne <> 0 And alertTimeTwo <> 0 And alertTimeFour <> 0 Then
Exit Sub
End If

myMsgBox = MsgBox("Warning, you are activating all timers!", vbOKCancel, "TIMERS WILL BE ACTIVATED")
If myMsgBox = vbCancel Then
Exit Sub
End If
MsgBox "You activated all timers successfully.", vbOKOnly, "Successfully Activated Timers"


Call manualTimerReset


ws.Range("C6").Value = Now

End Sub
