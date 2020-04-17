

Dim http As New WinHttp.WinHttpRequest
Public RunWhen As Double
Public NoMatch As Boolean
Public Timer1 As Long

Option Explicit

Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare PtrSafe Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Const IF_FROM_CACHE = &H1000000
Private Const IF_MAKE_PERSISTENT = &H2000000
Private Const IF_NO_CACHE_WRITE = &H4000000

Private Const BUFFER_LEN = 256
Dim source As String

Public Function GetURLSource(sURL As String)
    On Error GoTo ErrorHandler
    
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
    
    iResult = InternetCloseHandle(hInternet)
    
    GetURLSource = sData
    
    Exit Function
ErrorHandler:
    Resume Next
End Function


Function HttpExists(sURL As String) As Boolean
    On Error GoTo HttpError
    http.Open "GET", sURL
    http.Send
    HttpExists = (http.status = 200)
    
    Exit Function
HttpError:
    HttpExists = False
    Resume Next
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTelogisCSV
' Author    : Philip Deckers
' Date      : 8/20/2015
' Purpose   :
'---------------------------------------------------------------------------------------
Sub GetTelogisCSV(TemplateID As String, _
                  Organization As String, _
                  User As String, _
                  Pass As String, _
                  Optional Parameters As String = "")

Dim strURL As String
Dim strTDE As String
Dim iStartRow As Integer
Dim iStartCol As Integer
Dim iCurrentRow As Integer
Dim wb As Workbook
Dim ws As Worksheet
Dim myCSVRange As Range
Dim myCSVColumn As Range
Set wb = ThisWorkbook
Set ws = wb.Sheets("Data")
Set myCSVRange = ws.Range("A1")
Set myCSVColumn = ws.Range("A:A")

   On Error GoTo GetTelogisCSV_Error


'prepend ampersand to Parameters list if needed
If Parameters <> "" And Left(strTDE, 1) <> "&" Then Parameters = "&" & Parameters

'assemble URL
strURL = "https://" & Organization & ".api.telogis.com/execute" & _
                        "?template=" & TemplateID & _
                        "&user=" & User & _
                        "&pass=" & Pass & Parameters


'MsgBox (strURL)

'retrieve HTTP data as string
strTDE = GetURLSource(strURL)
                        
iStartRow = 1
iStartCol = 1

Application.ScreenUpdating = False

Cells(iStartRow, iStartCol).Select
iCurrentRow = iStartRow

'parse returned TDE data onto separate rows by <CR> break
Do While InStr(1, strTDE, Chr(13), vbTextCompare) > 1
    Cells(iCurrentRow, iStartCol).Formula = Replace(Replace(Trim(Left(strTDE, InStr(1, strTDE, Chr(13), vbTextCompare))), Chr(13), ""), Chr(10), "")
    strTDE = Trim(Right(strTDE, Len(strTDE) - InStr(1, strTDE, Chr(13), vbTextCompare)))
    iCurrentRow = iCurrentRow + 1
    Loop

'use TextToColumns function to split comma separated values into separate columns

Columns(iStartCol).Select
Selection.TextToColumns _
    Destination:=Cells(iStartRow, iStartCol), _
    DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, _
    ConsecutiveDelimiter:=False, _
    Comma:=True, _
    TrailingMinusNumbers:=True

'AutoFit all cells
Cells.Select
Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit

Cells(iStartRow, iStartCol).Select

Application.ScreenUpdating = True

'

   On Error GoTo 0
   Exit Sub

GetTelogisCSV_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetTelogisCSV of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : Philip Deckers
' Date      : 8/20/2015
' Purpose   :
'---------------------------------------------------------------------------------------

Sub Main()

   On Error GoTo Main_Error

Dim strTemplate As String
Dim strOrganization As String
Dim strUser As String
Dim strPass As String
Dim strParameters As String
Dim lowDateInputBox As String
Dim upperDateInputBox As String
Dim sName As Name
Dim wb As Workbook
Dim ws As Worksheet
Dim buttonsWS As Worksheet
Dim myRange As Range
Dim myDataRange As Range
Set wb = ThisWorkbook
Set ws = wb.Sheets("Data")
Set myRange = ws.Range("A:T")
Set buttonsWS = wb.Sheets("BUTTONS")
Set myDataRange = ws.Range("A1")


Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
'ws.Visible = xlSheetVisible

If telAdHoc = False Then
    On Error Resume Next
    Application.OnTime alertTimeOne, "Main", , False
    alertTimeOne = 0




 If telTimerTest <> "" Then
        alertTimeOne = Now + TimeValue(telTimerTest)
        Application.OnTime alertTimeOne, "Main", defTelSchedUpper
        buttonsWS.Range("C18").Value = telTimerTest
        buttonsWS.Range("C12").Value = alertTimeOne
    End If
    If telTimerTest = "" Then
        alertTimeOne = Now + TimeValue(defaultTel)
'        Application.OnTime alertTimeThree, "dashBoardEmail", defDashSchedUpper
        alertTimeOneAsNumber = alertTimeOne - Int(alertTimeOne)
        If alertTimeOneAsNumber > defTelSchedUpper Then
            Application.OnTime defTelSched, "Main", defTelSchedUpper
            buttonsWS.Range("C12").Value = defTelSched
        Else
        Application.OnTime alertTimeOne, "Main", defTelSchedUpper
        buttonsWS.Range("C18").Value = defaultTel
        buttonsWS.Range("C12").Value = alertTimeOne
        End If
    End If



For Each sName In ThisWorkbook.Names
    If InStr(1, sName, "Data") Then
    sName.Delete
    End If
Next

'lowDateInputBox = InputBox(Prompt:="Format as YYYY-MM-DD", Title:="Lower Date Input", Default:="Please enter a valid format LOWER date range")

    'display message box with value held by variable
    'MsgBox "Your input was: " & lowDateInputBox
'upperDateInputBox = InputBox(Prompt:="Format as YYYY-MM-DD", Title:="Upper Date Input", Default:="Please enter a valid format UPPER date range")

    'display message box with value held by variable
    'MsgBox "Your input was: " & upperDateInputBox



'strTemplate = "872983925"
''Parameters for template 872983925
'strParameters = "Start=2015-08-18T00:00:00z&End=2015-08-20T23:59:59z"


strTemplate = "1340134306" 'RetrieveJobActuals_MainesProd_CP01_CSV
''Parameters for template 1340134306
strParameters = "ReportStart=" & lowerDateRange() & "T00:00:00z&ReportEnd=" & upperDateRange() & "T23:59:59z"

strOrganization = "maines"
strUser = strOrganization & ":" & "apiuser"
strPass = "telogis!23"


'ActiveWorkbook.Sheets("Data").Select
'Cells.ClearContents
'Cells(1, 1).Select

ws.Activate
ws.AutoFilter.ShowAllData
myRange.Cells.ClearContents
Cells(1, 1).Select


GetTelogisCSV TemplateID:=strTemplate, Organization:=strOrganization, User:=strUser, Pass:=strPass, Parameters:=strParameters

   On Error GoTo 0
   


Call createNamedRange

buttonsWS.Activate
buttonsWS.Range("A3").Value = Now

' If telTimerTest <> "" Then
'        alertTimeOne = Now + TimeValue(telTimerTest)
'        Application.OnTime alertTimeOne, "Main", defTelSchedUpper
'        buttonsWS.Range("C18").Value = telTimerTest
'        buttonsWS.Range("C12").Value = alertTimeOne
'    End If
'    If telTimerTest = "" Then
'        alertTimeOne = Now + TimeValue(defaultTel)
''        Application.OnTime alertTimeThree, "dashBoardEmail", defDashSchedUpper
'        alertTimeOneAsNumber = alertTimeOne - Int(alertTimeOne)
'        If alertTimeOneAsNumber > defTelSchedUpper Then
'            Application.OnTime defTelSched, "Main", defTelSchedUpper
'            buttonsWS.Range("C12").Value = defTelSched
'        Else
'        Application.OnTime alertTimeOne, "Main", defTelSchedUpper
'        buttonsWS.Range("C18").Value = defaultTel
'        buttonsWS.Range("C12").Value = alertTimeOne
'        End If
'    End If


'alertTimeOne = Now + TimeValue("01:00:00")
'Application.OnTime alertTimeOne, "Main"
'
'buttonsWS.Range("C18").Value = alertTimeOne

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
''ws.Visible = xlSheetHidden

Exit Sub

Else

For Each sName In ThisWorkbook.Names
    If InStr(1, sName, "Data") Then
    sName.Delete
    End If
Next

'lowDateInputBox = InputBox(Prompt:="Format as YYYY-MM-DD", Title:="Lower Date Input", Default:="Please enter a valid format LOWER date range")

    'display message box with value held by variable
    'MsgBox "Your input was: " & lowDateInputBox
'upperDateInputBox = InputBox(Prompt:="Format as YYYY-MM-DD", Title:="Upper Date Input", Default:="Please enter a valid format UPPER date range")

    'display message box with value held by variable
    'MsgBox "Your input was: " & upperDateInputBox



'strTemplate = "872983925"
''Parameters for template 872983925
'strParameters = "Start=2015-08-18T00:00:00z&End=2015-08-20T23:59:59z"


strTemplate = "1340134306" 'RetrieveJobActuals_MainesProd_CP01_CSV
''Parameters for template 1340134306
strParameters = "ReportStart=" & lowerDateRange() & "T00:00:00z&ReportEnd=" & upperDateRange() & "T23:59:59z"

strOrganization = "maines"
strUser = strOrganization & ":" & "apiuser"
strPass = "telogis!23"


'ActiveWorkbook.Sheets("Data").Select
'Cells.ClearContents
'Cells(1, 1).Select

ws.Activate
ws.AutoFilter.ShowAllData
myRange.Cells.ClearContents
Cells(1, 1).Select


GetTelogisCSV TemplateID:=strTemplate, Organization:=strOrganization, User:=strUser, Pass:=strPass, Parameters:=strParameters

   On Error GoTo 0
   


Call createNamedRange

buttonsWS.Activate
buttonsWS.Range("A3").Value = Now


 If telTimerTest <> "" Then
        alertTimeOne = Now + TimeValue(telTimerTest)
        Application.OnTime alertTimeOne, "Main"
        buttonsWS.Range("C18").Value = telTimerTest
        buttonsWS.Range("C12").Value = alertTimeOne
    End If
    If telTimerTest = "" Then
        alertTimeOne = Now + TimeValue(defaultTel)
        Application.OnTime alertTimeOne, "Main"
        buttonsWS.Range("C18").Value = defaultTel
        buttonsWS.Range("C12").Value = alertTimeOne
    End If

End If
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
'ws.Visible = xlSheetHidden
Exit Sub
Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module Module1"
    
End Sub


Public Function lowerDateRange()

Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
Dim myDisplayUpper As Range
Dim myDisplayLower As Range
Set myDisplayUpper = ws.Range("C8")
Set myDisplayLower = ws.Range("C9")
Dim dayAsNumber As Integer
Dim dayAsNumberOG As Integer
Dim monthAsNumber As Integer
Dim yearAsNumber As Integer
Dim myDay As String
Dim myMonth As String
Dim myYear As String
Dim myYesterDay As String
Dim myYesterMonth As String
Dim myYesterYear As String
Dim myPassOne As String
Dim myPassTwo As String


Dim myRange1 As Range
Dim blank As String



blank = "0"

myDay = VBA.DateTime.Day(Date)
If Len(myDay) = 1 Then
    myDay = blank & myDay
End If

myMonth = VBA.DateTime.Month(Date)
If Len(myMonth) = 1 Then
    myMonth = blank & myMonth
End If

myYear = VBA.DateTime.Year(Date)
If Len(myYear) = 1 Then
    myYear = blank & myYear
End If

myPassOne = myYear & "-" & myMonth & "-" & myDay

dayAsNumberOG = VBA.DateTime.Day(Date)
dayAsNumber = VBA.DateTime.Day(Date)
    If dayAsNumber > 2 Then
        dayAsNumber = dayAsNumber - 2
        Else
            If dayAsNumber = 1 And ((myMonth = "05") Or (myMonth = "07") Or (myMonth = "10") Or (myMonth = "12")) Then
                dayAsNumber = "29"
            End If
            If dayAsNumber = 2 And ((myMonth = "05") Or (myMonth = "07") Or (myMonth = "10") Or (myMonth = "12")) Then
                dayAsNumber = "30"
            End If
            If dayAsNumber = 1 And ((myMonth = "01") Or (myMonth = "02") Or (myMonth = "04") Or (myMonth = "06") Or (myMonth = "08") Or (myMonth = "09") Or (myMonth = "11")) Then
                dayAsNumber = "30"
            End If
            If dayAsNumber = 2 And ((myMonth = "01") Or (myMonth = "02") Or (myMonth = "04") Or (myMonth = "06") Or (myMonth = "08") Or (myMonth = "09") Or (myMonth = "11")) Then
                dayAsNumber = "31"
            End If
            If dayAsNumber = 1 And myMonth = "03" Then
                dayAsNumber = "27"
            End If
            If dayAsNumber = 2 And myMonth = "03" Then
                dayAsNumber = "28"
            End If
    End If

myYesterDay = dayAsNumber
If Len(myYesterDay) = 1 Then
    myYesterDay = blank & myYesterDay
End If

monthAsNumber = VBA.DateTime.Month(Date)
    If ((dayAsNumberOG = 1) Or (dayAsNumberOG = 2)) And myMonth = "01" Then
        myYesterMonth = "12"
        monthAsNumber = myYesterMonth
    End If
        If ((dayAsNumberOG = 1) Or (dayAsNumberOG = 2)) And myMonth <> "01" Then
            monthAsNumber = monthAsNumber - 1
    End If
myYesterMonth = monthAsNumber
If Len(myYesterMonth) = 1 Then
    myYesterMonth = blank & myYesterMonth
End If

yearAsNumber = VBA.DateTime.Year(Date)
    If ((dayAsNumberOG = 1) Or (dayAsNumberOG = 2)) And myMonth = "01" Then
        yearAsNumber = yearAsNumber - 1
    End If
myYesterYear = yearAsNumber
If Len(myYesterYear) = 1 Then
    myYesterYear = blank & myYesterYear
End If


myPassTwo = myYesterYear & "-" & myYesterMonth & "-" & myYesterDay
myDisplayLower = myPassTwo

lowerDateRange = myPassTwo

End Function
Public Function upperDateRange()

Dim wb As Workbook
Dim ws As Worksheet
Set wb = ThisWorkbook
Set ws = wb.Sheets("BUTTONS")
Dim myDisplayUpper As Range
Dim myDisplayLower As Range
Set myDisplayUpper = ws.Range("C8")
Set myDisplayLower = ws.Range("C9")

Dim dayAsNumber As Integer
Dim myDay As String
Dim myMonth As String
Dim myYear As String
Dim myYesterDay As String
Dim myPassOne As String
Dim myPassTwo As String


Dim myRange1 As Range
Dim blank As String



blank = "0"

myDay = VBA.DateTime.Day(Date)
If Len(myDay) = 1 Then
    myDay = blank & myDay
End If

myMonth = VBA.DateTime.Month(Date)
If Len(myMonth) = 1 Then
    myMonth = blank & myMonth
End If

myYear = VBA.DateTime.Year(Date)
If Len(myYear) = 1 Then
    myYear = blank & myYear
End If

myPassOne = myYear & "-" & myMonth & "-" & myDay


dayAsNumber = VBA.DateTime.Day(Date)
If dayAsNumber = 1 And myMonth = "05" Or myMonth = "07" Or myMonth = "10" Or myMonth = "12" Then
    dayAsNumber = "30"
        If dayAsNumber = 1 And myMonth = "01" Or myMonth = "02" Or myMonth = "04" Or myMonth = "06" Or myMonth = "08" Or myMonth = "09" Or myMonth = "11" Then
            dayAsNumber = "31"
            End If
                If dayAsNumber = 1 And myMonth = "03" Then
                    dayAsNumber = "28"
                    End If
    Else
    dayAsNumber = dayAsNumber - 1
End If

myYesterDay = dayAsNumber
If Len(myYesterDay) = 1 Then
    myYesterDay = blank & myYesterDay
End If

myPassTwo = myYear & "-" & myMonth & "-" & myYesterDay
myDisplayUpper = myPassOne

upperDateRange = myPassOne

End Function
