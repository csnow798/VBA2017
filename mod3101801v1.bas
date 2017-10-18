Attribute VB_Name = "Ä£¿é2"
Sub mod3101801v1()
    Dim i, xrow As Integer
    xrow = 1
    For i = 1 To 100
        If i Mod 3 = 0 Then
            Cells(xrow, "B") = i
            xrow = xrow + 1
        End If
    Next
End Sub
