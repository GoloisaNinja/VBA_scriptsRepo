Sub JumptoTMcell()
'use this format to move about long data sets
'current format uses offset to shift cell 3 columns to the right for data entry

Dim sCell As String

sCell = InputBox( _
    Prompt:="TM Number?", _
    Title:="Input TM number please")
On Error GoTo ErrMsg

    Cells.Columns(1).Find(sCell).Select
    ActiveCell.offset(0, 3).Select
    Exit Sub
ErrMsg:
    MsgBox ("Whoopsie daisy, that was unexpected"), vbCritical, "Oh Darn"
End Sub

