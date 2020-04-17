Option Explicit
Public mySecDayArr() As Variant
Public destSheet As Variant
Public dest2Sheet As Variant

Sub dayTwoTrips()
Dim wb As Workbook
Dim ws As Worksheet
Dim buttonsWS As Worksheet
Dim routeWS As Worksheet
Dim lastRow As Long
Dim x As Long
Dim realLastRow As Long
Dim i, v
Dim rng As Range
Dim tbl As TableObject
Dim TempFilePath As String
Dim TempFileName As String
Dim WBB As Workbook
Dim targetWSheet As Worksheet
Dim sourceWSheet As Worksheet
Dim tableDestWSheet As Worksheet
Dim target2WSheet As Worksheet
Dim source2WSheet As Worksheet
Dim tableDest2WSheet As Worksheet
Dim shtName As String
Dim outapp As Object
Dim outmail As Object
Set wb = ThisWorkbook
Set ws = wb.Sheets("Route Summary")
Set routeWS = wb.Sheets("ROUTED BY ACCT")
ws.AutoFilter.ShowAllData
routeWS.AutoFilter.ShowAllData
Set buttonsWS = wb.Sheets("BUTTONS")
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Application.DisplayAlerts = False
Application.ScreenUpdating = False


For i = 3 To lastRow
        If ws.Range("A" & i).Value <> 0 Then
            realLastRow = realLastRow + 1
        End If
    Next i
realLastRow = realLastRow + 2
ReDim Preserve mySecDayArr(0)
For v = 3 To realLastRow
        If IsError(ws.Range("I" & v).Value) Then
            ws.Range("I" & v).Value = ""
        End If
        If Not ws.Range("I" & v).Value = "Complete" Then
            If UBound(mySecDayArr) >= 0 Then
            mySecDayArr(UBound(mySecDayArr)) = ws.Range("A" & v).Value
            ReDim Preserve mySecDayArr(0 To UBound(mySecDayArr) + 1)
            End If
        End If
    Next v
    
    Set outapp = CreateObject("Outlook.Application")
    Set outmail = outapp.CreateItem(0)
        On Error Resume Next
        With outmail
            .To = "jonathan.collins@maines.net"
            .CC = ""
            .BCC = ""
            .Subject = "Second Day Incomplete Route Data " & Now
            .HTMLBody = buildSecDayDataEmail()
            
                    Application.ScreenUpdating = False
                    TempFilePath = Environ$("temp") & "\"
                    TempFileName = "SecondDayRouteData" & " " & Format(Now, "dd-mmm-yy hh-mm-ss")
                    Set WBB = Workbooks.Add
                    With WBB
                    .SaveAs TempFilePath & TempFileName, FileFormat:=51
                    End With
                    Set targetWSheet = WBB.Worksheets("Sheet1")
                    Set sourceWSheet = wb.Sheets("secDayRoutes")
                    Set target2WSheet = WBB.Worksheets("Sheet2")
                    Set source2WSheet = wb.Sheets("secDayRoutesDep")
                    'wb.Sheets(lateTrip).Copy after:=targetWSheet
                    sourceWSheet.Copy After:=targetWSheet
                    Set tableDestWSheet = WBB.Worksheets("secDayRoutes")
                    Set rng = tableDestWSheet.Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
                    Set tbl = tableDestWSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
                    source2WSheet.Copy After:=WBB.Sheets("secDayRoutes")
                    Set tableDest2WSheet = WBB.Worksheets("secDayRoutesDep")
                    Set rng = tableDest2WSheet.Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
                    Set tbl = tableDest2WSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)

                    With WBB

                        .SaveAs TempFilePath & TempFileName, FileFormat:=51
                    End With


            .Attachments.Add WBB.FullName
            .Send
        End With
    WBB.Close savechanges:=False
    Kill TempFilePath & TempFileName & ".xlsx"
    wb.Sheets("secDayRoutes").Delete
    wb.Sheets("secDayRoutesDep").Delete
    
    Set outmail = Nothing
    Set outapp = Nothing
    
buttonsWS.Activate
buttonsWS.Range("P11").Value = Now

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Public Function buildSecDayDataEmail()
    Dim oSheet As Worksheet
    Dim pSheet As Worksheet
    Dim wb As Workbook
    Dim tbl As ListObject
    Dim rng As Range
    Dim tbl2 As ListObject
    Dim rng2 As Range
    Set wb = ThisWorkbook
    Set oSheet = wb.Sheets("ROUTED BY ACCT")
    Set pSheet = wb.Sheets("Routes with Departure")
    Dim i, lastRoww, n, lastRowww, c, d
    lastRoww = oSheet.Cells(Rows.Count, 1).End(xlUp).Row
    lastRowww = pSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim html, custRoute, custStop, custNum, custName, custAdd, custState, custCity, custPhone, custSchedTime, custCases, custRoadNet, custDriver, custDepot
    Dim depRoute, depRteName, depDepotNum, depDelvStops, depBackHaul, depCases, depMiles, depPlanDep, depRouteTime, depServiceTime, depTravelTime, depLayoverTime, dep1Driver, dep2Driver
    
    Dim lngDestLrow As Long

    Dim myRangeA As Range
    Dim myRangeB As Range
    Dim myRangeC As Range
    Dim myRangeD As Range
    Dim myRangeE As Range
    Dim myRangeF As Range
    Dim myRangeG As Range
    Dim myRangeH As Range
    Dim myRangeI As Range
    Dim myRangeJ As Range
    Dim myRangeK As Range
    Dim myRangeL As Range
    Dim myRangeM As Range
    Dim myMasterRange As Range

    wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).Name = "secDayRoutes"
    
    Dim lngDest2Lrow As Long

    Dim my2RangeA As Range
    Dim my2RangeB As Range
    Dim my2RangeC As Range
    Dim my2RangeD As Range
    Dim my2RangeE As Range
    Dim my2RangeF As Range
    Dim my2RangeG As Range
    Dim my2RangeH As Range
    Dim my2RangeI As Range
    Dim my2RangeJ As Range
    Dim my2RangeK As Range
    Dim my2RangeL As Range
    Dim my2RangeM As Range
    Dim my2RangeN As Range
    Dim my2MasterRange As Range

    wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)).Name = "secDayRoutesDep"
    
    'wb.Sheets(Sheets.Count).Name = lateTrip

    Set destSheet = wb.Sheets("secDayRoutes")
    Set dest2Sheet = wb.Sheets("secDayRoutesDep")

    Set myRangeA = destSheet.Range("A1")
    Set myRangeB = destSheet.Range("B1")
    Set myRangeC = destSheet.Range("C1")
    Set myRangeD = destSheet.Range("D1")
    Set myRangeE = destSheet.Range("E1")
    Set myRangeF = destSheet.Range("F1")
    Set myRangeG = destSheet.Range("G1")
    Set myRangeH = destSheet.Range("H1")
    Set myRangeI = destSheet.Range("I1")
    Set myRangeJ = destSheet.Range("J1")
    Set myRangeK = destSheet.Range("K1")
    Set myRangeL = destSheet.Range("L1")
    Set myRangeM = destSheet.Range("M1")
    
    Set myMasterRange = destSheet.Range("A:M")
    myMasterRange.NumberFormat = "@"
    
    
    Set my2RangeA = dest2Sheet.Range("A1")
    Set my2RangeB = dest2Sheet.Range("B1")
    Set my2RangeC = dest2Sheet.Range("C1")
    Set my2RangeD = dest2Sheet.Range("D1")
    Set my2RangeE = dest2Sheet.Range("E1")
    Set my2RangeF = dest2Sheet.Range("F1")
    Set my2RangeG = dest2Sheet.Range("G1")
    Set my2RangeH = dest2Sheet.Range("H1")
    Set my2RangeI = dest2Sheet.Range("I1")
    Set my2RangeJ = dest2Sheet.Range("J1")
    Set my2RangeK = dest2Sheet.Range("K1")
    Set my2RangeL = dest2Sheet.Range("L1")
    Set my2RangeM = dest2Sheet.Range("M1")
    Set my2RangeN = dest2Sheet.Range("N1")
    
    Set my2MasterRange = dest2Sheet.Range("A:N")
    my2MasterRange.NumberFormat = "@"

    myRangeA.Value = "Route"
    myRangeB.Value = "Stop"
    myRangeC.Value = "Customer"
    myRangeD.Value = "Customer Name"
    myRangeE.Value = "Address"
    myRangeF.Value = "City"
    myRangeG.Value = "State"
    myRangeH.Value = "Phone"
    myRangeI.Value = "Sched. Time"
    myRangeJ.Value = "Cases"
    myRangeK.Value = "Roadnet Planned Arrival"
    myRangeL.Value = "Driver Name"
    myRangeM.Value = "Depot"

    lngDestLrow = destSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    my2RangeA.Value = "Route"
    my2RangeB.Value = "Route Name"
    my2RangeC.Value = "Depot"
    my2RangeD.Value = "Delivery Stops"
    my2RangeE.Value = "Backhauls"
    my2RangeF.Value = "Cases"
    my2RangeG.Value = "Miles"
    my2RangeH.Value = "Plan Departure"
    my2RangeI.Value = "Route Time"
    my2RangeJ.Value = "Service Time"
    my2RangeK.Value = "Travel Time"
    my2RangeL.Value = "Layover Time"
    my2RangeM.Value = "Driver 1"
    my2RangeN.Value = "Driver 2"

    lngDest2Lrow = dest2Sheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    
    
    
    html = "<!DOCTYPE html><html><body>"
    html = html & "<div style=""font-family:Arial; font-size: 10px; max-width: 768px;"">"
    html = html & "<h2 style='text-align: left; font-family: Arial;'>" & "Please see the attached file containing incomplete 2nd day routes" & "</h2>"
    html = html & "<h4 style='text-align: left; font-family: Arial;'>" & "Copy and Paste the data from tab " & "<u>" & "secDayRoutes" & "</u>" & " into the ROUTED BY ACCT tab" & "</h4>"
    html = html & "<h4 style='text-align: left; font-family: Arial;'>" & "Copy and Paste the data from tab " & "<u>" & "secDayRoutesDep" & "</u>" & " into the Routes by Departure tab" & "</h4>"
    
 For i = LBound(mySecDayArr) To UBound(mySecDayArr)
 
    
    For n = 2 To lastRoww

        custRoute = oSheet.Range("A" & n).Value
        custStop = oSheet.Range("B" & n).Value
        custNum = oSheet.Range("C" & n).Value
        custName = oSheet.Range("D" & n).Value
        custAdd = oSheet.Range("E" & n).Value
        custCity = oSheet.Range("F" & n).Value
        custState = oSheet.Range("G" & n).Value
        custPhone = oSheet.Range("H" & n).Value
        custSchedTime = oSheet.Range("I" & n).Value
        custCases = oSheet.Range("J" & n).Value
        custRoadNet = oSheet.Range("K" & n).Value
        custDriver = oSheet.Range("L" & n).Value
        custDepot = oSheet.Range("M" & n).Value
        
        
        If mySecDayArr(i) = custRoute Then
        On Error Resume Next
        lngDestLrow = destSheet.Cells(Rows.Count, "A").End(xlUp).Row
        destSheet.Cells(lngDestLrow + 1, "A") = custRoute
        destSheet.Cells(lngDestLrow + 1, "B") = custStop
        destSheet.Cells(lngDestLrow + 1, "C") = custNum
        destSheet.Cells(lngDestLrow + 1, "D") = custName
        destSheet.Cells(lngDestLrow + 1, "E") = custAdd
        destSheet.Cells(lngDestLrow + 1, "F") = custCity
        destSheet.Cells(lngDestLrow + 1, "G") = custState
        destSheet.Cells(lngDestLrow + 1, "H") = custPhone
        destSheet.Cells(lngDestLrow + 1, "I") = custSchedTime
        destSheet.Cells(lngDestLrow + 1, "J") = custCases
        destSheet.Cells(lngDestLrow + 1, "K") = custRoadNet
        destSheet.Cells(lngDestLrow + 1, "L") = custDriver
        destSheet.Cells(lngDestLrow + 1, "M") = custDepot
        On Error Resume Next
        
        
        
    End If
    Next n
    
Next i
        
        
For c = LBound(mySecDayArr) To UBound(mySecDayArr)
 
    
    For d = 2 To lastRowww

        depRoute = pSheet.Range("A" & d).Value
        depRteName = pSheet.Range("B" & d).Value
        depDepotNum = pSheet.Range("C" & d).Value
        depDelvStops = pSheet.Range("D" & d).Value
        depBackHaul = pSheet.Range("E" & d).Value
        depCases = pSheet.Range("F" & d).Value
        depMiles = pSheet.Range("G" & d).Value
        depPlanDep = pSheet.Range("H" & d).Value
        depRouteTime = pSheet.Range("I" & d).Value
        depServiceTime = pSheet.Range("J" & d).Value
        depTravelTime = pSheet.Range("K" & d).Value
        depLayoverTime = pSheet.Range("L" & d).Value
        dep1Driver = pSheet.Range("M" & d).Value
        dep2Driver = pSheet.Range("N" & d).Value
        
       
        If mySecDayArr(c) = depRoute Then
        On Error Resume Next
        lngDest2Lrow = dest2Sheet.Cells(Rows.Count, "A").End(xlUp).Row
        dest2Sheet.Cells(lngDest2Lrow + 1, "A") = depRoute
        dest2Sheet.Cells(lngDest2Lrow + 1, "B") = depRteName
        dest2Sheet.Cells(lngDest2Lrow + 1, "C") = depDepotNum
        dest2Sheet.Cells(lngDest2Lrow + 1, "D") = depDelvStops
        dest2Sheet.Cells(lngDest2Lrow + 1, "E") = depBackHaul
        dest2Sheet.Cells(lngDest2Lrow + 1, "F") = depCases
        dest2Sheet.Cells(lngDest2Lrow + 1, "G") = depMiles
        dest2Sheet.Cells(lngDest2Lrow + 1, "H") = depPlanDep
        dest2Sheet.Cells(lngDest2Lrow + 1, "I") = depRouteTime
        dest2Sheet.Cells(lngDest2Lrow + 1, "J") = depServiceTime
        dest2Sheet.Cells(lngDest2Lrow + 1, "K") = depTravelTime
        dest2Sheet.Cells(lngDest2Lrow + 1, "L") = depLayoverTime
        dest2Sheet.Cells(lngDest2Lrow + 1, "M") = dep1Driver
        dest2Sheet.Cells(lngDest2Lrow + 1, "N") = dep2Driver
        On Error Resume Next
        
        
        
    End If
    Next d
    
Next c
        

    html = html & "</div></body></html>"
    buildSecDayDataEmail = html
    
    
'    Dim tbl As ListObject
'    Dim rng As Range

    Set rng = destSheet.Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = destSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.TableStyle = "TableStyleMedium15"
    
    Set rng2 = dest2Sheet.Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl2 = dest2Sheet.ListObjects.Add(xlSrcRange, rng2, , xlYes)
    tbl2.TableStyle = "TableStyleMedium15"



End Function
