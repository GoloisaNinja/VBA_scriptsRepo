Public Sub ResetButton(ByRef btn As Object)
'damn button keeps changing size and font and everyting else...i'm tired of it
'so boom...

Dim h As Integer    'command button height
Dim w As Integer    '               width
Dim fs As Integer   '               font size
    With btn
        h = 30
        w = 131.25
        fs = 11
        .AutoSize = True
        .AutoSize = False
        .Height = h
        .Width = w
        .Font.Size = fs
    End With
End Sub
