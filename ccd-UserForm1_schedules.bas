 Option Explicit

Public Sub CommandButton1_Click()
Dim wb As Workbook
Dim ws As Worksheet
Dim masterAlert As Double
Dim schedMaster As Double
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")

masterAlert = alertTimeOne + alertTimeTwo + alertTimeThree + alertTimeFour
schedMaster = defTelSched + defIntConSched + defDashSched

On Error Resume Next

If OptionButton1.Value = True Then
        If alertTimeOne <> 0 Or defTelSched <> 0 Or restoreTelSched <> 0 Then
            Application.OnTime alertTimeOne, "Main", , False
            alertTimeOne = 0
            Application.OnTime defTelSched, "Main", , False
            alertTimeOne = 0
            ws.Range("C12") = alertTimeOne
            MsgBox "Success", vbOKOnly, "Disabled Telogis Event"
        Else
            MsgBox "Event(s) were already disabled", vbOKOnly, "No action needed"
        End If
End If
If OptionButton2.Value = True Then
            If alertTimeTwo <> 0 Or defIntConSched <> 0 Or restoreIntConSched <> 0 Then
            Application.OnTime alertTimeTwo, "allRouteBlast", , False
            alertTimeTwo = 0
            Application.OnTime defIntConSched, "allRouteBlast", , False
            alertTimeTwo = 0
            ws.Range("C13") = alertTimeTwo
            MsgBox "Success", vbOKOnly, "Disabled Internal Consolidated Event"
        Else
            MsgBox "Event(s) were already disabled", vbOKOnly, "No action needed"
        End If
End If
If OptionButton3.Value = True Then
            If alertTimeFour <> 0 Then
                Application.OnTime alertTimeFour, "newDynamicRoute", , False
                alertTimeFour = 0
                ws.Range("C14") = alertTimeFour
                MsgBox "Success", vbOKOnly, "Disabled External Late Event"
            Else
                MsgBox "Event(s) were already disabled", vbOKOnly, "No action needed"
            End If
End If
If OptionButton4.Value = True Then
            
                Application.OnTime alertTimeThree, "executiveDash", , False
                alertTimeThree = 0
                ws.Range("C15") = alertTimeThree
                Application.OnTime defDashSched, "executiveDash", , False
'                dashAdHoc = True
'                ws.Range("G21").Value = dashAdHoc
                MsgBox "Success", vbOKOnly, "Disabled MGMT Dashboard Event"
           
End If

masterAlert = alertTimeOne + alertTimeTwo + alertTimeThree + alertTimeFour
If OptionButton5.Value = True Then
            If masterAlert <> 0 Or schedMaster <> 0 Then
                Call disableTimerEvents
            Else
                MsgBox "Event(s) were already disabled", vbOKOnly, "No action needed"
            End If
End If


ws.Range("A6") = Now
Unload Me

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Initialize()

End Sub
