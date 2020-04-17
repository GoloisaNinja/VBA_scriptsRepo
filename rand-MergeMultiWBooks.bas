Sub MergeMultiBooks()
Dim wbDst As Workbook
Dim wbSrc As Workbook
Dim wsSrc As Worksheet
Dim MyPath As String
Dim strFilename As String

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
MyPath = "C:\Users\Jonathan Collins\Documents\CFA TEST 2"
Set wbDst = Workbooks.Add(xlWBATWorksheet)
strFilename = Dir(MyPath & "\*.xlsx", vbNormal)

If Len(strFilename) = 0 Then Exit Sub

Do Until strFilename = ""

Set wbSrc = Workbooks.Open(Filename:=MyPath & "\" & strFilename)

Set wsSrc = wbSrc.Worksheets(1)

wsSrc.Copy After:=wbDst.Worksheets(wbDst.Worksheets.Count)

wbSrc.Close False

strFilename = Dir()

Loop
wbDst.Worksheets(1).Delete
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub
