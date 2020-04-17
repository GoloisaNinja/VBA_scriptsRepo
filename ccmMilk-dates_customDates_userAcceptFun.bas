Function RndInt(lowerbound As Integer, upperbound As Integer) As Integer
Randomize
RndInt = Int(lowerbound + Rnd() * (upperbound - lowerbound + 1))
End Function
Sub newMonth()
Dim myArr As Variant
Dim myFunArr As Variant
Dim wb As Workbook
Dim mainS As Worksheet
Dim myRange0 As Range
Dim myRange1 As Range
Dim myMonth As String
Dim myMonth1 As Integer
Dim myMonth2 As Integer
Dim rand1 As Integer
Dim check1 As Boolean
Dim check2 As Boolean
Dim check3 As Boolean
Dim myRange2 As Range
Dim myRange3 As Range
Dim strPass As String

strPass = "holidaygolo04"
Set wb = ThisWorkbook
Set mainS = wb.Sheets("MAIN")

Set myRange0 = mainS.Range("H74")
Set myRange1 = mainS.Range("I74")
Set myRange2 = mainS.Range("H76:H98")
Set myRange3 = mainS.Range("K76:K98")

'unprotect sheet so macro can run
mainS.Unprotect Password:=strPass
If Not Range("B6:B9").Locked = False Then
        mainS.Range("B6:B9").Locked = False
End If

'creating our arrays
myArr = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
myFunArr = Array("See those keys that aren't letters? Try hitting one of those.", "Are you feeling alright?", "Mushrooms can affect brain functions.", "Why you gotta be difficult?", "I'm not sure you should be using this tool.")

'create a boolean check to use in loop
check1 = False

While check1 = False
    myMonth = InputBox("Enter the whole number of the upcoming pricing month", "Nice to see you", "e.g. 1,2,3...")
        'first step checks for cancel box is user changes mind, exits sub
        If myMonth = "" Then
            mainS.Protect Password:=strPass, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, UserInterfaceOnly:=True
            Exit Sub
        End If
        'second step checks if input was numeric, if not a do loop keeps asking for a number with fun new random titles
        If Not IsNumeric(myMonth) Then
            Do While Not IsNumeric(myMonth)
            rand1 = RndInt(0, 4)
            MsgBox ("C'mon man....just enter a real month number. You can do it, I believe in you."), vbOKOnly, myFunArr(rand1)
            myMonth = InputBox("Enter the whole number of the upcoming pricing month", "Nice to see you", "e.g. 1,2,3...")
            Loop
        End If
        'third step checks if number input is a month number or not, ridicules you if it isn't
        If Int(myMonth) > 12 Or Int(myMonth) < 1 Then
            rand1 = RndInt(0, 4)
            MsgBox ("C'mon man....just enter a real month number. You can do it, I believe in you."), vbOKOnly, myFunArr(rand1)
        Else
        'finally if all goes well and the user isn't just trying to be painful then we change boolean of check1 and move on
            check1 = True
        End If
Wend

    If myMonth = 1 Then
        myMonth1 = 0
        myMonth2 = 11
    Else
        myMonth1 = myMonth - 1
        myMonth2 = myMonth - 2
    End If

myRange1 = myArr(myMonth1)
myRange0 = myArr(myMonth2)

myRange2.Value2 = myRange3.Value2

'reprotect after value change
mainS.Protect Password:=strPass, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, UserInterfaceOnly:=True

End Sub
