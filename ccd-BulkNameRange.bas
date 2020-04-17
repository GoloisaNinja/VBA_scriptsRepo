Sub createNamedRange()

Dim myWorksheet As Worksheet
Dim myNamedRange As Range
Dim myRangeName As String
 
Set myWorksheet = ThisWorkbook.Worksheets("Data")
Set myNamedRange1 = myWorksheet.Range("AC:AC")
Set myNamedRange2 = myWorksheet.Range("AD:AD")
Set myNamedRange3 = myWorksheet.Range("AF:AF")
Set myNamedRange4 = myWorksheet.Range("AH:AH")
Set myNamedRange5 = myWorksheet.Range("AI:AI")
Set myNamedRange6 = myWorksheet.Range("AP:AP")
Set myNamedRange7 = myWorksheet.Range("I:I")
Set myNamedRange8 = myWorksheet.Range("J:J")
Set myNamedRange9 = myWorksheet.Range("L:L")
Set myNamedRange10 = myWorksheet.Range("M:M")
Set myNamedRange11 = myWorksheet.Range("S:S")
Set myNamedRange12 = myWorksheet.Range("U:U")
Set myNamedRange13 = myWorksheet.Range("Y:Y")
Set myNamedRange14 = myWorksheet.Range("Z:Z")
Set myNamedRange15 = myWorksheet.Range("AQ:AQ")
Set myNamedRange16 = myWorksheet.Range("AR:AR")
Set myNamedRange17 = myWorksheet.Range("AS:AS")
Set myNamedRange18 = myWorksheet.Range("AT:AT")
Set myNamedRange19 = myWorksheet.Range("AU:AU")
Set myNamedRange20 = myWorksheet.Range("E:E")
Set myNamedRange21 = myWorksheet.Range("AV:AV")
Set myNamedRange22 = myWorksheet.Range("AN:AN")
Set myNamedRange23 = myWorksheet.Range("AW:AW")

myRangeName1 = "lateEarly"
myRangeName2 = "absTimeDiff"
myRangeName3 = "lastDel"
myRangeName4 = "totalRoutes"
myRangeName5 = "missedMark"
myRangeName6 = "dataStart"
myRangeName7 = "dataEarlyWin"
myRangeName8 = "dataLatestWin"
myRangeName9 = "dataActArrTime"
myRangeName10 = "JobStatus"
myRangeName11 = "dataEstArrTime"
myRangeName12 = "routeStop"
myRangeName13 = "combRtSt"
myRangeName14 = "dataPlanArr"
myRangeName15 = "ActArrHelper"
myRangeName16 = "EstArrHelper"
myRangeName17 = "AbsTimeDiffHelper"
myRangeName18 = "LateEarlyHelper"
myRangeName19 = "DataStartHelper"
myRangeName20 = "RouteStartTime"
myRangeName21 = "CombStartHelper"
myRangeName22 = "DEP_TIME"
myRangeName23 = "TwoDayHelp"


 ThisWorkbook.Names.Add Name:=myRangeName1, RefersTo:=myNamedRange1
 ThisWorkbook.Names.Add Name:=myRangeName2, RefersTo:=myNamedRange2
 ThisWorkbook.Names.Add Name:=myRangeName3, RefersTo:=myNamedRange3
 ThisWorkbook.Names.Add Name:=myRangeName4, RefersTo:=myNamedRange4
 ThisWorkbook.Names.Add Name:=myRangeName5, RefersTo:=myNamedRange5
 ThisWorkbook.Names.Add Name:=myRangeName6, RefersTo:=myNamedRange6
 ThisWorkbook.Names.Add Name:=myRangeName7, RefersTo:=myNamedRange7
 ThisWorkbook.Names.Add Name:=myRangeName8, RefersTo:=myNamedRange8
 ThisWorkbook.Names.Add Name:=myRangeName9, RefersTo:=myNamedRange9
 ThisWorkbook.Names.Add Name:=myRangeName10, RefersTo:=myNamedRange10
 ThisWorkbook.Names.Add Name:=myRangeName11, RefersTo:=myNamedRange11
 ThisWorkbook.Names.Add Name:=myRangeName12, RefersTo:=myNamedRange12
 ThisWorkbook.Names.Add Name:=myRangeName13, RefersTo:=myNamedRange13
 ThisWorkbook.Names.Add Name:=myRangeName14, RefersTo:=myNamedRange14
 ThisWorkbook.Names.Add Name:=myRangeName15, RefersTo:=myNamedRange15
 ThisWorkbook.Names.Add Name:=myRangeName16, RefersTo:=myNamedRange16
 ThisWorkbook.Names.Add Name:=myRangeName17, RefersTo:=myNamedRange17
 ThisWorkbook.Names.Add Name:=myRangeName18, RefersTo:=myNamedRange18
 ThisWorkbook.Names.Add Name:=myRangeName19, RefersTo:=myNamedRange19
 ThisWorkbook.Names.Add Name:=myRangeName20, RefersTo:=myNamedRange20
 ThisWorkbook.Names.Add Name:=myRangeName21, RefersTo:=myNamedRange21
 ThisWorkbook.Names.Add Name:=myRangeName22, RefersTo:=myNamedRange22
 ThisWorkbook.Names.Add Name:=myRangeName23, RefersTo:=myNamedRange23

End Sub

