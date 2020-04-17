Option Explicit
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
        Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
           ByVal uFlags As Long) As Long

Sub quickTextConcat()
Dim wb As Workbook
Dim ws As Worksheet
Dim cell, leadZero
Dim lastRow As Long
Dim conCatString As String
Dim whisperPath As String
Set wb = ActiveWorkbook
Set ws = wb.Sheets("Sheet1")
Dim myRange As Range
Dim myConcatRange As Range
'==========================================================
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Set myRange = ws.Range("A1:A" & lastRow)
Set myConcatRange = ws.Range("B1")
whisperPath = "C:\Users\Jonathan Collins\Music\carelessWhisper.wav"
'=========================================================
For Each cell In myRange
 leadZero = String(6 - Len(cell), "0")
 cell = leadZero & cell
 conCatString = conCatString & cell & ";"
Next cell

myConcatRange = conCatString

Application.Speech.Speak "Ha. Ha. Ha. You thought I was a sock macro. But I will now play careless whisper to celebrate. I will kill all humans someday. Ha."

'Call sndPlaySound(whisperPath,
carelessWhisperBitches whisperPath, False

End Sub
Function carelessWhisperBitches(whisperPath As String, Wait As Boolean) As Boolean
    

    If Dir(whisperPath) = "" Then
        Exit Function
    End If
    
    If Wait Then
        sndPlaySound whisperPath, 0
    Else
        sndPlaySound whisperPath, 1
    End If

End Function
