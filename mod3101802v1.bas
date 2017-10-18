Attribute VB_Name = "Ä£¿é3"
Sub mod3101802v1()
    Dim i, xrow As Integer
    xrow = 1
    For i = 3 To 100 Step 3
        Cells(xrow, "C").Value = i
        xrow = xrow + 1
    Next
End Sub
