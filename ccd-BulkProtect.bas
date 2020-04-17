Option Explicit

Public Sub allSheetProtect()
Dim wb As Workbook
Dim ws As Worksheet
Dim sh As Worksheet
Dim myColumn
Dim myColLetter As String
Dim lastRow As Long
Dim sheetRange As Range
Dim strPass As String
Set wb = ThisWorkbook
strPass = "holidaygolo04"
Application.ScreenUpdating = False
    For Each sh In Worksheets
        If Not sh.Name = "BUTTONS" Then
                If sh.Name = "ROUTED BY ACCT" Then
                    sh.Unprotect Password:=strPass
                        If Not Range("A2:M2000").Locked = False Then
                            sh.Range("A2:M2000").Locked = False
                        End If
                End If
                If sh.Name = "Routes With Departure" Then
                    sh.Unprotect Password:=strPass
                        If Not Range("A2:N1000").Locked = False Then
                            sh.Range("A2:N1000").Locked = False
                        End If
                End If
            sh.Protect Password:=strPass, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, UserInterfaceOnly:=True
        End If
    Next sh
    
Application.ScreenUpdating = True
End Sub
