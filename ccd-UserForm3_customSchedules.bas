Option Explicit

Public Sub CommandButton1_Click()
Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")



On Error Resume Next

    MsgBox "You are applying a custom timer interval with no scheduled stop time. " & Chr(10) & _
        "The adHoc state for this event will change to true." & Chr(10) & _
        "The custom timer interval will execute until disabled.", vbOKOnly, "Custom Event Handler"
        
If OptionButton1.Value = True Then
        If TextBox1.Value <> "" Then
            Application.OnTime defTelSched, "Main", , False
            defTelSched = TimeValue("00:00:00")
            defTelSchedUpper = TimeValue("00:00:00")
            ws.Range("G12").Value = defTelSched
            ws.Range("H12").Value = defTelSchedUpper
            telTimerTest = TextBox1.Value
            Application.OnTime alertTimeOne, "Main", , False
            alertTimeOne = 0
            alertTimeOne = Now + TimeValue(telTimerTest)
            Application.OnTime alertTimeOne, "Main"
            ws.Range("C12") = alertTimeOne
            ws.Range("C18") = telTimerTest
            telAdHoc = True
            ws.Range("G18").Value = telAdHoc
        End If
        If TextBox1.Value = "" Then
            MsgBox "No custom value entered", vbOKOnly, "No changes made"
        End If
    End If


If OptionButton2.Value = True Then
        If TextBox1.Value <> "" Then
            Application.OnTime defIntConSched, "allRouteBlast", , False
            defIntConSched = TimeValue("00:00:00")
            defIntConSchedUpper = TimeValue("00:00:00")
            ws.Range("G13").Value = defIntConSched
            ws.Range("H13").Value = defIntConSchedUpper
            intConTimerTest = TextBox1.Value
            Application.OnTime alertTimeTwo, "allRouteBlast", , False
            alertTimeTwo = 0
            alertTimeTwo = Now + TimeValue(intConTimerTest)
            Application.OnTime alertTimeTwo, "allRouteBlast"
            ws.Range("C13") = alertTimeTwo
            ws.Range("C19") = intConTimerTest
            intConAdHoc = True
            ws.Range("G19").Value = intConAdHoc
        End If
        If TextBox1.Value = "" Then
            MsgBox "No custom value entered", vbOKOnly, "No changes made"
        End If
    End If


    If OptionButton4.Value = True Then
        If TextBox1.Value <> "" Then
            Application.OnTime defDashSched, "executiveDash", , False
            defDashSched = TimeValue("00:00:00")
            defDashSchedUpper = TimeValue("00:00:00")
            ws.Range("G15").Value = defDashSched
            ws.Range("H15").Value = defDashSchedUpper
            customTimerTest = TextBox1.Value
            Application.OnTime alertTimeThree, "executiveDash", , False
            alertTimeThree = 0
            alertTimeThree = Now + TimeValue(customTimerTest)
            Application.OnTime alertTimeThree, "executiveDash"
            ws.Range("C15") = alertTimeThree
            ws.Range("C21") = customTimerTest
            dashAdHoc = True
            ws.Range("G21").Value = dashAdHoc
        End If
        If TextBox1.Value = "" Then
            MsgBox "No custom value entered", vbOKOnly, "No changes made"
        End If
    End If


ws.Range("R5").Value = Now
Unload Me

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Initialize()

End Sub

