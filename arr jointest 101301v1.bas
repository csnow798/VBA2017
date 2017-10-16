Attribute VB_Name = "Ä£¿é3"
Sub JoinTest()
    Dim arr As Variant, txt As String
    arr = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    txt = Join(arr, "@")
    MsgBox txt
End Sub
