'==========================================================
'NEW MATERIAL SETUP EDR CAPTURE CODE
'VER 1.1
'AUTHOR: JON COLLINS (VISTEX) 2019-2020
'=========================================================

'IT IS EXTREMELY IMPORTANT THIS CODE BE PLACED WITHIN - THISWORKBOOK - OF THE MICROSOFT EXCEL OBJECT WITHIN WHICH YOU WANT IT TO FIRE
'===============================================================================================================================================
Option Explicit
'Const VK_VOLUME_MUTE = &HAD
'Const VK_VOLUME_DOWN = &HAE
'Const VK_VOLUME_UP = &HAF
'Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'================================================================================================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)

  Dim ws As Worksheet
  Dim wb As Workbook
  Dim outapp As Object
  Dim outmail As Object
  Dim matAnswer As String
  Dim chkExit As Boolean
  Dim matRet
  Dim thorCheck As Integer
  Dim matDesc As String
  Dim useName As String
  Dim samGroup, julGroup, ryanGroup As String
  Dim myNewDict As Object
  Set myNewDict = CreateObject("Scripting.Dictionary")
  myNewDict.CompareMode = vbTextCompare
  Dim subbD As String
  Dim subbE As String
  Set wb = ThisWorkbook
  Set ws = wb.Sheets("ITEM FORM")

'INITIALLY SETTING THIS FALSE - WILL USE LATER AS A CANCEL CHECK AND EARLY EXIT FROM SUB CONDITION

  chkExit = False
  matDesc = ws.Range("H14").Value

'SEE USERNAME FUNCTION (LINE 128) - THIS FUNCTION SAFELY CHECKS USER NAME OF COMPUTER TO BE USED AS CONDITIONAL LATER ON

  useName = UserName()

'IF THE SHEET IS CLOSED BY ANYONE NOT EQUAL TO THE NAMES BELOW - THE SUBROUTINE EXITS IMMEDIATELY

  If useName <> "Jonathan Collins" Then
      Exit Sub
  End If

'THESE ARE NOT THE DROIDS YOU ARE LOOKING FORM

'  Call minVol
'  Call rVol

'SOME BASIC SHEET VALIDATION

    If ws.Range("AC18").Value = "" Then
        MsgBox "EDR/PROGRAM field must be set to YES or NO", vbOKOnly, "STOP EVERYTHING"
        Cancel = True
        Exit Sub
    End If
    
    If ws.Range("AC18").Value = "YES" And ws.Range("G12").Value = "" Then
        If ws.Range("AA12").Value = "" Then
            MsgBox "EDR/PROGRAM field is set to YES, but no EDR information exists - Please address", vbOKOnly, "STOP EVERYTHING"
            Cancel = True
            Exit Sub
        End If
    End If
    
    
    If ws.Range("AE4").Value = "Y" And ws.Range("H6").Value = "" Then
        MsgBox "Cost Contracted is Y but Sales Mat Grp 1 is empty - Please address", vbOKOnly, "STOP EVERYTHING"
        Cancel = True
        Exit Sub
    End If
    

'SETTING MATANSWER VAR EQUAL TO MATANSWERCHECK FUNTION (LINE 147) - WHICH TAKES TWO PARAMS - MATRET (ANTICIPATING MAT NUM RETURN) AND CHKEXIT (WHICH IS INITIALLY SET TO FALSE)

  matAnswer = matAnswerCheck(matRet, chkExit)

'IF CHKEXIT COMES BACK FROM MATANSWERCHECK AS TRUE THEN THE USER CANCELLED ON PURPOSE AND WE EXIT THE SUB

  If chkExit Then
      
      Exit Sub
  End If

'IF WE PASS THE CHKEXIT IFSTATEMENT THEN WE APPLY THE MATRET VALUE GOTTEN IN THE MATANSWERCHECK FUNTION TO VAR MATANSWER TO BE PASSED IN LATER FUNCTIONS

  matAnswer = matRet
  
'EMAIL COLLECTION DATA STRUCTURE BASED ON SUBMISSION FIELD

 samGroup = "holidayninjastaff@gmail.com; jonathan.collins@live.com"
 julGroup = "myInternalTest"
 
 myNewDict.Add key:="SAM ABBOTT", Item:=samGroup
 myNewDict.Add key:="JULIE REESE", Item:=julGroup
 myNewDict.Add key:="COLBAN MCGILLICUTTY III", Item:=samGroup
 
subbD = ws.Range("H4").Value
subbD = UCase(subbD)

If myNewDict.Exists(subbD) Then
subbE = myNewDict(subbD)
Else
subbD = Replace(subbD, " ", ".")
subbD = subbD & "@live.com; jonathan.collins@maines.net"
subbE = subbD
End If
'=========================================================
  
  Call sendySendy(subbE, "purch", matAnswer, matDesc, useName)
  
  
'INSTANTLY CHECK EDR FIELD INFORMATION - IF BLANK EXIT THE SUBROUTINE IMMEDIATELY

  If ws.Range("AC18").Value = "YES" Then
      Call sendySendy("jonathan.collins@maines.net", "edr", matAnswer, matDesc, useName)
  End If
  

End Sub
'WE DO THE SENDY SENDY ON CALLING THE OUTLOOK OBJECT

Public Function sendySendy(who As String, emailType As String, matAnswer, matDesc, useName)
    Dim emSub As String
    Dim typeS As String
    Dim sFile As String
    Dim TempFilePath As String
    Dim sAttach
    Dim outapp As Object
    Dim outmail As Object
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
    
    If emailType = "edr" Then
        typeS = "edr"
        emSub = "newItemEDR-TEST"
        emailType = BuildHtmlBody_main(matAnswer, typeS, useName)
          
          On Error Resume Next
                With outmail
                    .To = who
                    .CC = ""
                    .BCC = ""
                    .Subject = emSub
                    .HTMLBody = emailType
                    .Send
                End With
            Set outmail = Nothing
            Set outapp = Nothing
        
        Else
        typeS = "purch"
        emSub = matDesc & " assigned to " & matAnswer & " - " & Now()
        emailType = BuildHtmlBody_main(matAnswer, typeS, useName)
'=============================================================================================
'ADDED SECTION TO CREATE COPY OF ACTIVEWORK BOOK AND STORE IT IN A TEMP ENVIRONMENT FOLDER
'TEMP FILE IS KILLED BEFORE THE END OF THE FUNCTION
'THIS ALLOWS FOR ACTIVEWORKBOOK DATA TO BE ATTACHED AS EXCEL FILE SIMILAR WITH CURRENT PRACTICE
'==============================================================================================

        sFile = matDesc & " assigned to " & matAnswer & ".xlsm"
        TempFilePath = Environ$("temp") & "\"
        ActiveWorkbook.SaveCopyAs Filename:=TempFilePath & sFile
        
'===============================================================================================
          On Error Resume Next
          With outmail
              .To = who
              .CC = ""
              .BCC = ""
              .Subject = emSub
              .HTMLBody = emailType
              .Attachments.Add (TempFilePath & sFile)
              .Send
          End With
      Kill TempFilePath & sFile & ".xlsm"
      Set outmail = Nothing
      Set outapp = Nothing
      
    End If
End Function

'FUNCTION THAT HANDLES THE CONSTRUCTION OF THE EMAIL BODY TO BE PARSED AT A LATER TIME, ACCEPTS MATANSWER AS ARG

Public Function BuildHtmlBody_main(matAnswer, typeS As String, useName)

    Dim oSheet As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ITEM FORM")
    Dim html, nIDate, nIComp, nISub, nIVen, nIVenName, nIEDR As String, nIEDRexcl, nIMat, nIUOM, nIMatDesc, nIContract
    
    

    'GRABS VALUES OF ALL RELEVANT RANGES

    nIDate = Format(oSheet.Range("H2").Value, "mm/dd/yyyy")
    nIComp = oSheet.Range("AB2").Value
    nISub = oSheet.Range("H4").Value
    nIVen = oSheet.Range("AE8").Value
    nIVenName = oSheet.Range("H8").Value
    nIEDR = oSheet.Range("G12").Value
    nIEDRexcl = oSheet.Range("AA12").Value
    nIUOM = oSheet.Range("AG30").Value
    nIMatDesc = oSheet.Range("H14").Value
    nIContract = oSheet.Range("AE4").Value
    nIMat = matAnswer
    
    If nIUOM = "Y" Then
            nIUOM = "LB"
        Else
            nIUOM = "CS"
    End If

    'BUILDS THE HTML'

    If typeS = "edr" Then

    html = "<!DOCTYPE html><html><body>"
    html = html & "^"
    html = html & nIDate & "^"
    html = html & nIComp & "^"
    html = html & nISub & "^"
    html = html & nIVen & "^"
    html = html & nIVenName & "^"
    html = html & nIEDR & "^"
    html = html & nIEDRexcl & "^"
    html = html & nIMat & "^"
    html = html & nIUOM & "^"
    html = html & nIMatDesc & "^"
    html = html & nIContract
    html = html & "^"
    html = html & "</body></html>"

    'SETS AND RETURNS BUILDHTMLBODY EQUAL TO THE HTML JUST CONSTRUCTED ABOVE

    BuildHtmlBody_main = html
    
    Else
    
    html = "<!DOCTYPE html><html><body>"
    html = html & "<div>"
    html = html & "<h3 style='font-family: Arial;'>" & "The following material was setup" & "</h3>"
    html = html & "</div>"
    html = html & "<div style='font-family: Arial; font-size: 11px'>"
    html = html & "<table style='border-collapse: collapse; border-spacing: 0px; border-style: solid; border-color: #ccc; font-family: Arial'>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Date" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIDate & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Company" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIComp & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Submitted by" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nISub & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Vendor" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIVen & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Vendor Name" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIVenName & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Material" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIMat & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Material Desc" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & nIMatDesc & "</td></tr>"
    html = html & "<tr>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & "Setup By" & "</td>"
    html = html & "<td style='font-size: 11px; padding: 5px; border-style: solid; border-width: 1px 1px 0 0;'>" & useName & "</td></tr>"
    html = html & "</table>"
    html = html & "</div>"
    html = html & "</body></html>"

    'SETS AND RETURNS BUILDHTMLBODY EQUAL TO THE HTML JUST CONSTRUCTED ABOVE

    BuildHtmlBody_main = html
    End If
    

End Function

'FUNCTION TO SAFELY GRAB USERNAME FROM HOST COMPUTER

Public Function UserName(Optional WithDomain As Boolean = False) As String

    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")

    If WithDomain Then
        UserName = objNetwork.UserDomain & "\" & objNetwork.UserName
    Else
        UserName = objNetwork.UserName
    End If

    Set objNetwork = Nothing

End Function

'MATERIAL ENTRY FUNCTION THAT DOES SEVERAL CHECKS AGAINST CONSTRAINTS, ACCEPTS MATRET AND CHKEXIT ARGS

Public Function matAnswerCheck(matRet, chkExit)
    Dim strPattern As String
    Dim errorCheck As Integer
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    strPattern = "^[0-9]+$"
    regEx.Pattern = strPattern

'ERRORS RETURN USER TO THIS SECTION FOR REENTRY OF MATERIAL IF THAT PATH RESOLVES

LineMain:
    errorCheck = 0

'INPUTBOX COLLECTS MATERIAL NUMBER

    matRet = InputBox("Enter the material number")

'THIS SECTION CHECKS FOR MATERIAL INPUT BOX CANCELLATION AND REDIRECTS OR EXITS FUNCTION BASED ON RESPONSE

    If StrPtr(matRet) = 0 Then
            chkExit = True
            Exit Function
    End If
        


'CHECKS IF MATERIAL IS EMPTY, NOT EQUAL TO LENGHT OF SIX, AND AGAINST REGEX STATEMENT ABOVE THAT ONLY ACCEPTS STRING STARTING AND ENDING WITH DIGITS, DIGITS OF 0-9

    If matRet = "" Then
        errorCheck = errorCheck + 1
    End If

    If Len(matRet) <> 6 Then
        errorCheck = errorCheck + 1
    End If

    If Not regEx.Test(matRet) Then
        errorCheck = errorCheck + 1
    End If
    Debug.Print regEx.Test(matRet)

    Do While errorCheck > 0
        MsgBox "Not a valid material number...try again", vbOKCancel
        GoTo LineMain
    Loop


End Function

'================= THIS IS EXACTLY WHAT IT LOOKS LIKE - WE HIJACK USER COMPUTER VOL CONTROL AND UNMUTE/SET VOL TO 30 PERCENT ==========================
'DONT ASK QUESTIONS AND JUST ENJOY THE AWESOMENESS OF THIS

'Sub VolUp()
'   keybd_event VK_VOLUME_UP, 0, 1, 0
'   keybd_event VK_VOLUME_UP, 0, 3, 0
'End Sub
'
'Sub VolDown()
'   keybd_event VK_VOLUME_DOWN, 0, 1, 0
'   keybd_event VK_VOLUME_DOWN, 0, 3, 0
'End Sub
'
'Sub rVol()
'  Dim i As Integer
'    For i = 1 To 15
'      Call VolUp
'      Next i
'End Sub
'
'Sub minVol()
'  Dim i As Integer
'    For i = 1 To 100
'      Call VolDown
'      Next i
'End Sub
'
'
'Private Sub Workbook_Open()
'
'End Sub
