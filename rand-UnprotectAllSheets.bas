Sub Unprotect_All_Sheets()
Dim wSheet As Worksheet
Dim pwd As String
pwd = InputBox("Enter your password to unprotect all sheets", "Password Input")
On Error Resume Next
For Each wSheet In Worksheets
    wSheet.Unprotect Password:=pwd
Next wSheet
If Err <> 0 Then
MsgBox "You have entered an incorrect password.", vbCritical, "Oops, that was unexpected"
End If
On Error GoTo 0
End Sub
