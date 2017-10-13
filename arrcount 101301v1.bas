Attribute VB_Name = "Ä£¿é2"
Sub arrcount()
    Dim arr(10 To 50)
    MsgBox "Max Index Number is " & UBound(arr) & Chr(13) _
    & "Min Index Number is " & LBound(arr) & Chr(13) _
    & "Number of elements is " & UBound(arr) - LBound(arr) + 1
End Sub
