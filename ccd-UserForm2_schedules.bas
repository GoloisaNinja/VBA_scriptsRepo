 Option Explicit

Public Sub CommandButton1_Click()
Dim wb As Workbook
Dim ws As Worksheet
Dim masterAlert As Double
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")



On Error Resume Next
'If alertTimeOne <> 0 And alertTimeTwo <> 0 And alertTimeThree <> 0 And alertTimeFour <> 0 Then GoTo endGame
'If OptionButton1.Value = True And ((alertTimeOne = 0) Or (defTelSched = 0)) Then
'    alertTimeOne = Now + TimeValue(defaultTel)
'    Application.OnTime alertTimeOne, "Main"
'    ws.Range("C18") = alertTimeOne
'    MsgBox "Success", vbOKOnly, "Enabled Telogis Event"
'End If
If OptionButton1.Value = True Then
    telTimerTest = ""
    Application.OnTime alertTimeOne, "Main", , False
    alertTimeOne = 0
    alertTimeOne = Now + TimeValue(defaultTel)
    
    alertTimeOneAsNumber = alertTimeOne - Int(alertTimeOne)
    
    If alertTimeOneAsNumber < defTelSched Or alertTimeOneAsNumber > defTelSchedUpper Then
         MsgBox "Default event/schedule/timer have been reset and will run at the designated time in the event status board. If you need this event to run sooner, you'll need to set a custom timer.", vbOKOnly, "Event exceeds Schedule"
        alertTimeOne = 0
        telAdHoc = False
        defTelSched = restoreTelSched
        defTelSchedUpper = restoreTelSchedUpper
        ws.Range("G12") = defTelSched
        ws.Range("H12") = defTelSchedUpper
        ws.Range("G18") = telAdHoc
        Application.OnTime defTelSched, "Main", , False
        Application.OnTime defTelSched, "Main", defTelSchedUpper
        ws.Range("C12") = defTelSched
        ws.Range("C18") = defaultTel
    Else
        Application.OnTime alertTimeOne, "Main"
        telAdHoc = False
        defTelSched = restoreTelSched
        defTelSchedUpper = restoreTelSchedUpper
        ws.Range("G12") = defTelSched
        ws.Range("H12") = defTelSchedUpper
        ws.Range("G18") = telAdHoc
        ws.Range("C12") = alertTimeOne
        ws.Range("C18") = defaultTel
        MsgBox "Success", vbOKOnly, "Enabled Default Telogis Event and Schedule"
    End If
End If


If OptionButton2.Value = True Then
    intConTimerTest = ""
    Application.OnTime alertTimeTwo, "allRouteBlast", , False
    alertTimeTwo = 0
    alertTimeTwo = Now + TimeValue(defaultIntCon)
    
    alertTimeTwoAsNumber = alertTimeTwo - Int(alertTimeTwo)
    
    If alertTimeTwoAsNumber < defIntConSched Or alertTimeTwoAsNumber > defIntConSchedUpper Then
         MsgBox "Default event/schedule/timer have been reset and will run at the designated time in the event status board. If you need this event to run sooner, you'll need to set a custom timer.", vbOKOnly, "Event exceeds Schedule"
        alertTimeTwo = 0
        intConAdHoc = False
        defIntConSched = restoreIntConSched
        defIntConSchedUpper = restoreIntConSchedUpper
        ws.Range("G13") = defIntConSched
        ws.Range("H13") = defIntConSchedUpper
        ws.Range("G19") = intConAdHoc
        Application.OnTime defIntConSched, "allRouteBlast", , False
        Application.OnTime defIntConSched, "allRouteBlast", defIntConSchedUpper
        ws.Range("C13") = defIntConSched
        ws.Range("C19") = defaultIntCon
    Else
        Application.OnTime alertTimeTwo, "allRouteBlast"
        intConAdHoc = False
        defIntConSched = restoreIntConSched
        defIntConSchedUpper = restoreIntConSchedUpper
        ws.Range("G13") = defIntConSched
        ws.Range("H13") = defIntConSchedUpper
        ws.Range("G19") = intConAdHoc
        ws.Range("C13") = alertTimeTwo
        ws.Range("C19") = defaultIntCon
        MsgBox "Success", vbOKOnly, "Enabled Default Internal Consolidated Event and Schedule"
    End If
End If

If OptionButton3.Value = True And alertTimeFour = 0 Then
    alertTimeFour = Now + TimeValue("00:05:00")
    Application.OnTime alertTimeFour, "newDynamicRoute"
    ws.Range("C14") = alertTimeFour
    MsgBox "Success", vbOKOnly, "Enabled External Late Event"
End If

If OptionButton4.Value = True Then
    customTimerTest = ""
    Application.OnTime alertTimeThree, "excutiveDash", , False
    alertTimeThree = 0
    alertTimeThree = Now + TimeValue(defaultDash)
    
    alertTimeThreeAsNumber = alertTimeThree - Int(alertTimeThree)
    
    If alertTimeThreeAsNumber < defDashSched Or alertTimeThreeAsNumber > defDashSchedUpper Then
        MsgBox "Default event/schedule/timer have been reset and will run at the designated time in the event status board. If you need this event to run sooner, you'll need to set a custom timer.", vbOKOnly, "Event exceeds Schedule"
        alertTimeThree = 0
        dashAdHoc = False
        defDashSched = restoreDashSched
        defDashSchedUpper = restoreDashSchedUpper
        ws.Range("G15") = defDashSched
        ws.Range("H15") = defDashSchedUpper
        ws.Range("G21") = dashAdHoc
        Application.OnTime defDashSched, "executiveDash", , False
        Application.OnTime defDashSched, "executiveDash", defDashSchedUpper
        ws.Range("C15") = defDashSched
        ws.Range("C21") = defaultDash
    Else
        Application.OnTime alertTimeThree, "executiveDash"
        dashAdHoc = False
        defDashSched = restoreDashSched
        defDashSchedUpper = restoreDashSchedUpper
        ws.Range("G15") = defDashSched
        ws.Range("H15") = defDashSchedUpper
        ws.Range("G21") = dashAdHoc
        ws.Range("C15") = alertTimeThree
        ws.Range("C21") = defaultDash
        MsgBox "Success", vbOKOnly, "Enabled Default MGMT Dashboard Event and Schedule"
    End If
End If
If OptionButton5.Value = True Then
    Call reEnableTimerEvents
End If
ws.Range("D6") = Now
Unload Me
Exit Sub

endGame:
MsgBox "All Event(s) were already enabled", vbOKOnly, "No action needed"





Unload Me
Exit Sub
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Initialize()

End Sub

