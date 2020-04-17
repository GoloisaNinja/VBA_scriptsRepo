Sub edrJSON()
Dim wb As Workbook
Dim sh As Worksheet
Dim outapp As Object
Dim outmail As Object
Set outapp = CreateObject("Outlook.Application")
Set outmail = outapp.CreateItem(0)
Dim myEffDate As String
Dim mySubDate As String
Dim subBy As String
Dim jsonString As String
Dim newJsonString As String
Dim proceed101 As Integer
Dim venString, matString, edrString As String
Set wb = ThisWorkbook
Set ws = wb.Sheets("EDRFORM")
jsonString = "{" + """items""" + ":["
myEffDate = FormatDateTime(ws.Range("B3"), vbGeneralDate)
mySubDate = FormatDateTime(ws.Range("B2"), vbGeneralDate)
subBy = ws.Range("D2")

For i = 1 To Range("edrTable").Rows.Count

    venString = Format(ws.Range("edrTable[VENDOR]")(i))
    matString = Format(ws.Range("edrTable[MATERIAL]")(i), "@@@@@@")
    edrString = CStr(ws.Range("edrTable[EDR]")(i))
    
    jsonString = jsonString + "{" + """SUBDATE""" + ":" + """" + mySubDate + """" + "," + vbNewLine
    jsonString = jsonString + """EFFDATE""" + ":" + """" + myEffDate + """" + "," + vbNewLine
    jsonString = jsonString + """VEN#""" + ":" + """" + venString + """" + "," + vbNewLine
    jsonString = jsonString + """VENNAME""" + ":" + """" + ws.Range("edrTable[VENDORNAME]")(i) + """" + "," + vbNewLine
    jsonString = jsonString + """MAT""" + ":" + """" + matString + """" + "," + vbNewLine
    jsonString = jsonString + """MATDESC""" + ":" + """" + ws.Range("edrTable[MATERIALDESC]")(i) + """" + "," + vbNewLine
    jsonString = jsonString + """UOM""" + ":" + """" + ws.Range("edrTable[UOM]")(i) + """" + "," + vbNewLine
    jsonString = jsonString + """SUBBY""" + ":" + """" + subBy + """" + "," + vbNewLine
    jsonString = jsonString + """EDR""" + ":" + """" + edrString + """" + "},"
    

Next i

newJsonString = Left(jsonString, Len(jsonString) - 1)
newJsonString = newJsonString & "]}"

proceed101 = MsgBox("Do you really want to commit to sending this garbage?", vbOKCancel, "Slow ur roll homez")


If StrPtr(proceed101) = 0 Then
    Exit Sub
End If


          On Error Resume Next
                With outmail
                    .To = "jonathan.collins@maines.net"
                    .CC = ""
                    .BCC = ""
                    .Subject = "edrUpdate-Doc"
                    .HTMLBody = newJsonString
                    .Send
                End With
            Set outmail = Nothing
            Set outapp = Nothing
            
MsgBox "EDR Update Form Submitted", vbOKOnly
             
End Sub
