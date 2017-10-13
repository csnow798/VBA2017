Attribute VB_Name = "Ä£¿é1"
Sub dtsz()
    Dim arr() As String
    Dim n As Long
    n = Application.WorksheetFunction.CountA(Range("A:A"))
    ReDim arr(1 To n) As String
    MsgBox "Ubound of arr() is " & n
End Sub
