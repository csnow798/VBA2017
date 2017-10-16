Attribute VB_Name = "Ä£¿é4"
Sub ArrToRng()
    Dim arr As Variant
    arr = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Range("F1:F9").Value = Application.WorksheetFunction.Transpose(arr)
End Sub
