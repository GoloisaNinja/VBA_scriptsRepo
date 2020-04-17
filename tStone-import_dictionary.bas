Sub importSheets()

'================================================================================================
'// TURNING STONE AUTOMATION - PRICING MACRO
'// WRITTEN BY JCOLLINS - 04.14.2020 YEAR OF THE RONA
'// MAIN SECTION LOOPS THRU POWERAUTOMATE GENERATED ONEDRIVE FILES
'// TO EXTRACT ALL SHEETS FOR SIMPLE MATCH WITH SAP ZPCLIST VARIANT
'================================================================================================


'================================================================================================
'/ GENERAL VARIABLE DECLARATION FOR ALL SECTIONS
'================================================================================================
Dim directory As String
Dim fileName As String
Dim newShName As String
Dim total As Integer
Dim srcWbk As Workbook
Dim srcWsh As Worksheet
Dim wb As Workbook
Dim ws As Worksheet
Dim wsZ As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("convQty")
Set wsZ = wb.Sheets("zsp1")
Dim lastRow As Long
Dim lastRoww As Long
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastRoww = wsZ.Cells(Rows.Count, 9).End(xlUp).Row

'===============================================================================================
'// THIS SECTION IS A BIT WONKY TO DEBUG/STEP THRU WITH THE SCREEN UPDATING TURNED OFF
'// LOOP GRABS SHEET1 OF EVERY FILE ENDING IN SOME FORM OF .XL - NOTE SPECIAL CHARS
'// EACH WORKBOOK IS OPENED/COPIED/PASTED/CLOSED IN SUCCESSION
'===============================================================================================

Application.ScreenUpdating = False
Application.DisplayAlerts = False

directory = "C:\Users\Jonathan Collins\OneDrive - Maines Paper & Food Service\Documents\tsTemplates\"
fileName = Dir(directory & "*.xl??")

Do While fileName <> ""

Set srcWbk = Workbooks.Open(directory & fileName)
Set srcWsh = srcWbk.Sheets(1)

    
    total = ThisWorkbook.Worksheets.Count
    srcWsh.Copy After:=ThisWorkbook.Worksheets(total)
    



Workbooks(fileName).Close
fileName = Dir()

Loop



'================================================================================================
'// SECTION TWO
'// THIS SECTION WILL CREATE A DICTIONARY - USING TS ITEM CODES AND MARC'S STATIC CONV FACTORS
'// THE DICTIONARY BUILDS OFF THE CONVQTY SHEET EVERY RUN - SO THIS MUST BE UPDATED TO
'// TO MAINTAIN ACCURATE ZSP1 CONVERSION
'// THE FIRST FOR LOOP CRAFTS THE DICTIONARY - THE SECOND RUNS THE DICTIONARY AGAINST THE ZSP1
'// OUTPUTS
'=================================================================================================

Dim cell, keyI, itemC, zsp1, zcalc, zItem, eaCheck
Dim myRange As Range
Dim myNewDict As Object
Set myNewDict = CreateObject("Scripting.Dictionary")
myNewDict.CompareMode = vbTextCompare
Set myRange = ws.Range("A1:A" & lastRow)

ThisWorkbook.Save

For i = 2 To lastRow
    keyI = ws.Range("B" & i).Value
    itemC = ws.Range("E" & i).Value
        
    myNewDict.Add key:=keyI, Item:=itemC
            
    
Next i

For i = 2 To lastRoww
    zItem = wsZ.Range("A" & i).Value
    eaCheck = wsZ.Range("G" & i).Value
    If myNewDict.exists(zItem) Then
        If eaCheck <> "EA" Then
            wsZ.Range("I" & i).Value = wsZ.Range("I" & i).Value / myNewDict(zItem)
        End If
    End If
Next i


'========================================================================================================
'// SECTION THREE - MATCH IN THE MODIFIED ZSP1 VALUES TO THE VARIOUS SHEETS ALONG WITH
'// MAINES MATERIAL NUMBERS
'========================================================================================================

Dim tsKey, mainesItem, tsMatchKey, leadZero
Dim dynamicLastRow As Long
Dim loopSheet As Worksheet
Dim lsName As String
Dim tsDict As Object
Set tsDict = CreateObject("Scripting.Dictionary")
tsDict.CompareMode = vbTextCompare
Dim zspDict As Object
Set zspDict = CreateObject("Scripting.Dictionary")
zspDict.CompareMode = vbTextCompare


For i = 2 To lastRoww

    eaCheck = wsZ.Range("G" & i).Value
    
    If wsZ.Range("A" & i).Value <> "" Then
        If eaCheck <> "EA" Then
        
            tsKey = wsZ.Range("A" & i).Value
            mainesItem = wsZ.Range("B" & i).Value
            zspItem = wsZ.Range("I" & i).Value
        

                If Not tsDict.exists(tsKey) Then
                
                    tsDict.Add key:=tsKey, Item:=mainesItem
                    zspDict.Add key:=tsKey, Item:=zspItem
                    
                End If
        End If
        
    End If

Next i

For Each loopSheet In wb.Worksheets
    loopSheet.Activate
    lsName = ActiveSheet.Name
        If lsName <> "convQty" Then

            If lsName <> "zsp1" Then

                dynamicLastRow = loopSheet.Cells(Rows.Count, 1).End(xlUp).Row

                For i = 2 To dynamicLastRow
                    On Error Resume Next
                    tsMatchKey = Trim(Str(loopSheet.Range("A" & i).Value))
                    leadZero = String(6 - Len(tsMatchKey), "0")
                    tsMatchKey = leadZero & tsMatchKey
                    If tsDict.exists(tsMatchKey) Then
                        loopSheet.Range("E" & i).Value = tsDict(tsMatchKey)
                        loopSheet.Range("I" & i).Value = zspDict(tsMatchKey)
                    End If
                Next i
            End If
        End If
Next loopSheet

'==========================================================================================================
'// SECTION FOUR - COMPILE AND SEND AN EMAIL BACK TO MARC WITH EACH OF THE SHEETS AS A SEPARATE ATTACHMENT
'// THIS SECTION CREATES A TEMP FILE AND ATTACHES IT - THEN THE TEMP FILE IS DELETED FROM THE TEMP PATH
'// NOTE THE LOOP USED TO ATTACH THE MULTIPLE TEMP FILES
'==========================================================================================================

Dim WS_count As Integer
WS_count = wb.Worksheets.Count
Dim wbObj As Workbook
Dim currentWS As Worksheet
Dim srcAtSheet As Worksheet
Dim targetAtSheet As Worksheet
Dim outapp As Object
Dim outmail As Object
Dim sheetName As String
Dim TempFilePath As String
Dim TempFileName As String

Application.ScreenUpdating = False

Set outapp = CreateObject("Outlook.Application")
Set outmail = outapp.CreateItem(0)
On Error Resume Next
    With outmail
        .To = "jonathan.collins@maines.net"
        .CC = ""
        .BCC = ""
        .Subject = "Turning Stone Templates " & Now
        .HTMLBody = "<p>Please see attached templates</p>"

            For Each currentWS In wb.Worksheets
                currentWS.Activate
                sheetName = ActiveSheet.Name
                If sheetName <> "convQty" Then
                    If sheetName <> "zsp1" Then

                    TempFilePath = Environ$("temp") & "\"
                    TempFileName = sheetName & " " & Format(Now, "dd-mmm-yy hh-mm-ss")
                    Application.ScreenUpdating = False
                    Application.DisplayAlerts = False
                    Set wbObj = Workbooks.Add
                    With wbObj
                    .SaveAs TempFilePath & TempFileName, FileFormat:=51
                    End With
                    Set targetAtSheet = wbObj.Worksheets("Sheet1")
                    Set srcAtSheet = wb.Sheets(sheetName)
                    srcAtSheet.Copy After:=targetAtSheet

                    With wbObj

                        .SaveAs TempFilePath & TempFileName, FileFormat:=51

                    End With


                    .Attachments.Add wbObj.FullName

                    wbObj.Close savechanges:=False
                    Kill TempFilePath & TempFileName & ".xlsx"

                    End If
                End If

            Next currentWS

        .Send
    End With



Set outmail = Nothing
Set outapp = Nothing

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
