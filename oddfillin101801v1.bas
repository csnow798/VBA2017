Attribute VB_Name = "ģ��1"
Sub oddfillin()
    Dim i, xrow As Integer
    xrow = 1
    For i = 1 To 99 Step 2
        Cells(xrow, "A").Value = i
        xrow = xrow + 1
    Next
End Sub
